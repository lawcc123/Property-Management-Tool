"""Campaign SMTP protocol tests. All transports are fake; no network or real credentials."""
import hashlib
import base64
import io
import json
from pathlib import Path
import smtplib
import sys
import tempfile
import unittest
from unittest.mock import patch
sys.path.insert(0, str(Path(__file__).resolve().parents[1]/"PropertyManager.App"/"Backend"))
import campaign_mail as mail

class FakeSMTP:
    def __init__(self,*args,**kwargs):
        self.sent=[]
        self.failure=None
    def ehlo(self): pass
    def starttls(self,**kwargs): self.tls=kwargs['context']
    def login(self,user,password): self.login_values=(user,password)
    def quit(self): pass
    def send_message(self,message,**kwargs):
        self.sent.append((message,kwargs))
        if self.failure: raise self.failure
        return {}

class CampaignMailTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        file=Path(self.temp.name)/'Notice.pdf';file.write_bytes(b'%PDF-fake-fixture')
        self.message=dict(Id='message-1',Unit='201',OriginalTo='resident@example.test',To='test@example.test',Subject='[TEST - Unit 201] Notice',Body='Real fixture data',Attachments=[dict(Path=str(file),Name='201 - Notice.pdf',Hash=hashlib.sha256(file.read_bytes()).hexdigest().upper())])
        self.packet=dict(campaign=dict(Sender=dict(Server='smtp.example.test',Port=587,From='office@example.test'),Mode='Test',TestInbox='test@example.test',Messages=[self.message]),credentials=dict(username='login@example.test',password='fixture-secret'))
        self.smtp=FakeSMTP();self.events=[]
    def emit(self,kind,**data): self.events.append(dict(type=kind,**data))
    def run_worker(self,commands):
        return mail.run(self.packet,self.emit,io.StringIO(''.join(json.dumps(c)+'\n' for c in commands)),lambda *a,**k:self.smtp)
    def commands(self): return [dict(action='send',id='message-1'),dict(action='commit',id='message-1'),dict(action='quit')]
    def test_encoded_message_size_is_checked_before_submission_ack(self):
        with patch.object(mail, 'MESSAGE_LIMIT_BYTES', 300):
            self.assertEqual(self.run_worker(self.commands()), 1)
        self.assertFalse(self.smtp.sent)
        self.assertFalse(any(e['type'] == 'ready_to_submit' for e in self.events))
        self.assertTrue(any('20 MB' in e.get('message', '') for e in self.events))

    def test_raw_attachment_size_is_rejected_without_reading_content(self):
        with patch.object(mail, 'MESSAGE_LIMIT_BYTES', 5):
            self.assertEqual(self.run_worker(self.commands()), 1)
        self.assertFalse(self.smtp.sent)
        self.assertFalse(any(e['type'] == 'ready_to_submit' for e in self.events))

    def test_html_signature_has_plain_fallback_and_embedded_cid_picture(self):
        data = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aIu0AAAAASUVORK5CYII=')
        uri = 'data:image/png;base64,' + base64.b64encode(data).decode()
        self.message['HtmlBody'] = '<!DOCTYPE html><html><body><p>Dear resident</p><p><strong>Management Office</strong></p><img src="' + uri + '" /><img src="' + uri + '" /></body></html>'
        self.assertEqual(self.run_worker(self.commands()), 0)
        message = self.smtp.sent[0][0]
        html = message.get_body(preferencelist=('html',)).get_content()
        self.assertIn('cid:signature-', html)
        self.assertNotIn('data:image', html)
        self.assertIn(self.message['Body'], message.get_body(preferencelist=('plain',)).get_content())
        images = [p for p in message.walk() if p.get_content_type() == 'image/png']
        self.assertEqual(len(images), 1)
        self.assertEqual(images[0].get_content_disposition(), 'inline')
        self.assertEqual(images[0].get_payload(decode=True), data)
        self.assertEqual(self.smtp.sent[0][1]['to_addrs'], ['test@example.test'])

    def test_remote_signature_picture_blocks_before_ack(self):
        self.message['HtmlBody'] = '<html><body><img src="https://example.test/image.png" /></body></html>'
        self.assertEqual(self.run_worker(self.commands()), 1)
        self.assertFalse(self.smtp.sent)
        self.assertFalse(any(e['type'] == 'ready_to_submit' for e in self.events))

    def test_active_signature_html_blocks_before_ack(self):
        self.message['HtmlBody'] = '<html><body><script>alert(1)</script></body></html>'
        self.assertEqual(self.run_worker(self.commands()), 1)
        self.assertFalse(self.smtp.sent)

    def test_selected_real_data_sends_to_test_inbox_only_with_attachment(self):
        self.assertEqual(self.run_worker(self.commands()),0)
        message,kwargs=self.smtp.sent[0]
        self.assertEqual(kwargs['to_addrs'],['test@example.test'])
        self.assertEqual(message['To'],'test@example.test')
        self.assertIsNone(message['Cc']);self.assertIsNone(message['Bcc'])
        self.assertEqual(message['From'],'office@example.test')
        self.assertEqual(next(message.iter_attachments()).get_filename(),'201 - Notice.pdf')
        self.assertEqual([e['type'] for e in self.events],['ready','ready_to_submit','accepted'])
        self.assertEqual(self.smtp.login_values,('login@example.test','fixture-secret'))
    def test_no_journal_ack_means_no_submission(self):
        self.assertNotEqual(self.run_worker([dict(action='send',id='message-1'),dict(action='quit')]),0)
        self.assertEqual(self.smtp.sent,[])
    def test_test_destination_cannot_be_changed_to_resident(self):
        self.message['To']='resident@example.test'
        self.assertNotEqual(self.run_worker(self.commands()),0);self.assertEqual(self.smtp.sent,[])
    def test_live_destination_must_match_frozen_original(self):
        self.packet['campaign']['Mode']='Live'
        self.assertNotEqual(self.run_worker(self.commands()),0);self.assertEqual(self.smtp.sent,[])
    def test_six_test_messages_rejected(self):
        self.packet['campaign']['Messages']=[self.message]*6
        self.assertNotEqual(self.run_worker(self.commands()),0);self.assertEqual(self.smtp.sent,[])
    def test_changed_attachment_blocks_send_before_ack(self):
        Path(self.message['Attachments'][0]['Path']).write_bytes(b'changed')
        self.assertNotEqual(self.run_worker(self.commands()),0);self.assertEqual(self.smtp.sent,[])
    def test_data_rejection_is_definite_failure(self):
        self.smtp.failure=smtplib.SMTPDataError(550,b'rejected')
        self.run_worker(self.commands());self.assertFalse(self.events[-1]['uncertain'])
    def test_connection_loss_after_ack_is_uncertain(self):
        self.smtp.failure=smtplib.SMTPServerDisconnected('lost')
        self.run_worker(self.commands());self.assertTrue(self.events[-1]['uncertain'])
    def test_credentials_redacted(self):
        self.smtp.failure=RuntimeError('fixture-secret login@example.test')
        self.run_worker(self.commands());text=json.dumps(self.events)
        self.assertNotIn('fixture-secret',text);self.assertNotIn('login@example.test',text)
    def test_one_session_cannot_submit_same_message_twice(self):
        self.run_worker(self.commands()[:-1]+self.commands())
        self.assertEqual(len(self.smtp.sent),1)

if __name__=='__main__': unittest.main()
