"""SMTP-only interactive worker. Destinations are isolated and submission needs a journal ACK."""
import hashlib
import base64
from html.parser import HTMLParser
import json
import mimetypes
import smtplib
import ssl
import sys
from email.message import EmailMessage
from email.policy import SMTP

MESSAGE_LIMIT_BYTES = 20_000_000
from email.utils import parseaddr
from pathlib import Path


def valid_address(value):
    return isinstance(value, str) and "\r" not in value and "\n" not in value and "@" in value and parseaddr(value)[1] == value


def html_email(html):
    """Use only embedded PNG images; no external picture downloads or active content."""
    if len(html.encode('utf-8')) > MESSAGE_LIMIT_BYTES:
        raise ValueError("HTML email exceeds the 20 MB message budget")
    images = {}
    class Inspect(HTMLParser):
        def handle_starttag(self, tag, attrs):
            if tag not in {'html', 'head', 'meta', 'title', 'body', 'div', 'p', 'span', 'strong', 'b', 'em', 'i', 'u', 'br', 'a', 'img'}:
                raise ValueError("Unsupported HTML email content")
            values = dict(attrs)
            for key, value in attrs:
                if key.startswith('on') or key not in {'style', 'src', 'alt', 'width', 'height', 'href', 'charset'}:
                    raise ValueError("Unsupported HTML email attribute")
                if key == 'style' and any(x in (value or '').casefold() for x in ('url', 'expression', '<')):
                    raise ValueError("Unsupported HTML email style")
            if tag == 'a' and not values.get('href', '').lower().startswith(('https://', 'http://', 'mailto:')):
                raise ValueError("Unsupported signature link")
            if tag == 'img':
                uri = values.get('src', '')
                prefix = 'data:image/png;base64,'
                if not uri.startswith(prefix):
                    raise ValueError("Signature pictures must be embedded PNG images")
                data = base64.b64decode(uri[len(prefix):], validate=True)
                if len(data) > 1_000_000 or not data.startswith(b'\x89PNG\r\n\x1a\n'):
                    raise ValueError("Invalid or oversized signature picture")
                images[uri] = ('signature-' + hashlib.sha256(data).hexdigest()[:24], data)
    parser = Inspect(convert_charrefs=True)
    parser.feed(html)
    parser.close()
    for uri, (cid, _) in images.items():
        html = html.replace(uri, 'cid:' + cid)
    return html, list(images.values())


def compose(campaign, item):
    sender = campaign["Sender"]["From"]
    target = item["To"]
    if not valid_address(sender) or not valid_address(target):
        raise ValueError("A sender or recipient address is invalid")
    if campaign["Mode"] == "Test":
        if target.casefold() != campaign["TestInbox"].casefold() or not item["Subject"].startswith("[TEST - Unit "):
            raise ValueError("Test messages must go only to the test inbox")
    elif campaign["Mode"] != "Live" or target != item["OriginalTo"]:
        raise ValueError("The prepared live destination changed")
    message = EmailMessage()
    message["From"], message["To"], message["Subject"] = sender, target, item["Subject"]
    message["Message-ID"] = "<" + item["Id"] + "@" + sender.split("@", 1)[1] + ">"
    message.set_content(item["Body"])
    if item.get('HtmlBody'):
        html, images = html_email(item['HtmlBody'])
        message.add_alternative(html, subtype='html')
        html_part = message.get_payload()[-1]
        for cid, data in images:
            html_part.add_related(data, maintype='image', subtype='png', cid='<' + cid + '>', disposition='inline', filename='signature.png')
    raw_size = sum(Path(a["Path"]).stat().st_size for a in item["Attachments"])
    if raw_size > MESSAGE_LIMIT_BYTES:
        raise ValueError("Attachments exceed the 20 MB message budget; reduce files before sending")
    for attachment in item["Attachments"]:
        content = Path(attachment["Path"]).read_bytes()
        if hashlib.sha256(content).hexdigest().upper() != attachment["Hash"]:
            raise ValueError("A prepared attachment changed; prepare again before sending")
        kind = mimetypes.guess_type(attachment["Name"])[0] or "application/octet-stream"
        main, sub = kind.split("/", 1)
        message.add_attachment(content, maintype=main, subtype=sub, filename=attachment["Name"])
    if len(message.as_bytes(policy=SMTP)) > MESSAGE_LIMIT_BYTES:
        raise ValueError("The encoded email exceeds the 20 MB message budget; reduce files or message text")
    return message


def run(packet, emit, input_stream=None, transport=None):
    source = input_stream or sys.stdin
    smtp = None
    secret = packet.get("credentials", {})
    def error_text(error):
        text = str(error)
        for value in secret.values():
            if isinstance(value, str) and value:
                text = text.replace(value, "[redacted]")
        return text
    try:
        campaign = packet["campaign"]
        if campaign["Mode"] == "Test" and (not 1 <= len(campaign["Messages"]) <= 5 or not valid_address(campaign["TestInbox"])):
            raise ValueError("Choose 1 to 5 test messages and one valid test inbox")
        settings = campaign["Sender"]
        smtp = (transport or smtplib.SMTP)(settings["Server"], settings["Port"], timeout=30)
        smtp.ehlo()
        smtp.starttls(context=ssl.create_default_context())
        smtp.ehlo()
        smtp.login(secret["username"], secret["password"])
        emit("ready")
        allowed = {m["Id"]:m for m in campaign["Messages"]}
        for line in source:
            command = json.loads(line)
            if command.get("action") == "quit":
                break
            item = allowed.pop(command["id"])
            message = compose(campaign, item)
            emit("ready_to_submit", id=item["Id"])
            ack = json.loads(source.readline())
            if ack != {"action":"commit", "id":item["Id"]}:
                raise ValueError("No matching journal acknowledgement; email was not submitted")
            try:
                refused = smtp.send_message(message, to_addrs=[item["To"]])
                if refused:
                    emit("failed", id=item["Id"], uncertain=False, message="Recipient rejected by the mail server")
                    return 1
                emit("accepted", id=item["Id"])
            except (smtplib.SMTPRecipientsRefused, smtplib.SMTPSenderRefused, smtplib.SMTPDataError) as error:
                emit("failed", id=item["Id"], uncertain=False, message=error_text(error))
                return 1
            except Exception as error:
                emit("failed", id=item["Id"], uncertain=True, message=error_text(error))
                return 1
        return 0
    except Exception as error:
        emit("error", message=error_text(error))
        return 1
    finally:
        if smtp is not None:
            try:
                smtp.quit()
            except Exception:
                pass
