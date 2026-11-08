using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
namespace PropertyManager.App.PersonalizedEmail;
public partial class PersonalizedEmailView
{
    private readonly CampaignDeliveryStore deliveryStore;
    private readonly Func<ICampaignTransport> deliveryTransport;
    private PreparedCampaign? preparedLive;
    private readonly Dictionary<int,DeliveryMessage> preparedAttachments=[];
    private CancellationTokenSource? deliveryStop;
    private bool deliveryBusy;
    internal bool IsBusy=>deliveryBusy;
    internal void Cancel(){deliveryStop?.Cancel();StatusText.Text="Stopping after the current operation. Messages already accepted cannot be recalled.";}
    private void InvalidatePrepared(){preparedLive=null;preparedAttachments.Clear();if(SendLiveButton is not null)SendLiveButton.IsEnabled=false;if(LiveProgressPanel is not null)LiveProgressPanel.Visibility=Visibility.Collapsed;if(LiveSummary is not null)LiveSummary.Text="Messages changed. Continue through Review & test to prepare current live messages.";}
    private void DeliveryBusy(bool value)
    {
        deliveryBusy=value;EmailAccountSettingsButton.IsEnabled=!value;NewDraftButton.IsEnabled=OpenDraftButton.IsEnabled=SaveDraftButton.IsEnabled=!value;Steps.IsEnabled=HistoryNavigationButton.IsEnabled=NextButton.IsEnabled=BackButton.IsEnabled=DeliveryHistoryPanel.IsEnabled=!value;
        StopDeliveryButton.Visibility=value?Visibility.Visible:Visibility.Collapsed;DemoProgress.Visibility=value?Visibility.Visible:Visibility.Collapsed;
        if(value){StopDeliveryButton.Content="Stop after current operation";deliveryStop=new();DemoProgress.Value=0;}else{deliveryStop?.Dispose();deliveryStop=null;UpdateSteps();SendLiveButton.IsEnabled=preparedLive is not null;}
        UpdateSteps();
    }
    private IProgress<(string Message,int Done,int Total)> DeliveryProgress()=>new Progress<(string Message,int Done,int Total)>(p=>{StatusText.Text=p.Message+" ("+p.Done+" / "+p.Total+")";DemoProgress.Value=p.Total==0?0:100.0*p.Done/p.Total;});
    private async Task<PreparedCampaign> PrepareCampaign(IReadOnlyList<int>? selected,string? inbox)
    {
        var draft=ReadDraft();var sender=CampaignSender.Current();var progress=DeliveryProgress();var token=deliveryStop!.Token;
        return await Envelopes.Sta.Run(()=>new CampaignPreparation(deliveryStore,()=>new WordCampaignDocuments()).Prepare(draft,sender,selected,inbox,(m,d,t)=>progress.Report((m,d,t)),token));
    }
    private async Task PrepareLiveForSending()
    {
        if(deliveryBusy)return;
        ValidateStage(3);
        try
        {
            DeliveryBusy(true);
            var sender=CampaignSender.Current();
            if(preparedLive is null||CampaignPreparation.Fingerprint(ReadDraft(),sender)!=preparedLive.Fingerprint)preparedLive=await PrepareCampaign(null,null);
            CampaignPreparation.VerifyForDelivery(preparedLive,deliveryStore);
            preparedAttachments.Clear();for(int i=0;i<preparedLive.Messages.Count;i++)preparedAttachments[i]=preparedLive.Messages[i];ShowPreparedAttachment();
            var rows=deliveryStore.Rows(preparedLive);LiveSummary.Text=$"From: {preparedLive.Sender.From}\n{preparedLive.Messages.Count} live messages prepared. {rows.Count(r=>r.State=="Accepted")} already accepted; {rows.Count(r=>r.State=="Uncertain")} uncertain (both skipped). Testing is optional. No emails have been sent by preparation.";
            ReviewSummary.Text=LiveSummary.Text;StatusText.Text="Live messages prepared. Review the summary, then click Send live emails. A confirmation is required.";
        }
        catch{InvalidatePrepared();LiveSummary.Text="Preparation failed. Resolve the error, then choose Prepare / refresh live messages.";throw;}
        finally{DeliveryBusy(false);}
    }
    private async void PrepareAll_Click(object sender,RoutedEventArgs e)
    {
        try{await PrepareLiveForSending();}catch(Exception ex){Error(ex);}
    }
    private async void SendTests_Click(object sender,RoutedEventArgs e)
    {
        if(deliveryBusy)return;
        var selected=PreviewList.SelectedItems.Cast<CampaignPreview>().Select(p=>PreviewList.Items.IndexOf(p)).Order().ToList();string inbox=TestInbox.Text.Trim();
        if(selected.Count<1||selected.Count>5||!EmailLists.GeneratedEmailListService.IsValidEmail(inbox)){Warn("Select 1 to 5 message previews and enter one valid test inbox address.");return;}
        if(MessageBox.Show(OwnerWindow,$"Send {selected.Count} test message(s) only to {inbox}?\n\nThese use the selected residents' real data and generated attachments. No resident addresses will be used as delivery destinations. Live-send status will stay unchanged.","Send selected test emails",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return;
        try{await SendSelectedTests(selected,inbox);}catch(Exception ex){Error(ex);}
    }
    private async Task SendSelectedTests(IReadOnlyList<int> selected,string inbox)
    {
        try
        {
            DeliveryBusy(true);var campaign=await PrepareCampaign(selected,inbox);preparedAttachments.Clear();for(int i=0;i<selected.Count;i++)preparedAttachments[selected[i]]=campaign.Messages[i];ShowPreparedAttachment();
            await Deliver(campaign);
        }
        finally{DeliveryBusy(false);}
    }
    private bool ConfirmDelivery(PreparedCampaign campaign)
    {
        var rows=deliveryStore.Rows(campaign);int eligible=rows.Count(r=>r.State is "Not attempted" or "Failed");
        if(eligible==0){StatusText.Text="No eligible messages remain. Accepted and uncertain messages are not resent automatically. Check View history.";return false;}
        return MessageBox.Show(OwnerWindow,$"Send {eligible} {(campaign.Mode=="Test"?"TEST":"LIVE")} message(s)?\n\nFrom: {campaign.Sender.From}\n"+(campaign.Mode=="Test"?"Only destination: "+campaign.TestInbox:"Destinations: the prepared resident addresses")+$"\nAlready accepted: {rows.Count(r=>r.State=="Accepted")} (skipped)\nUncertain: {rows.Count(r=>r.State=="Uncertain")} (skipped)\n\nThis submits the prepared messages and attachments to the mail server.","Confirm email submission",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)==MessageBoxResult.Yes;
    }
    private async Task Deliver(PreparedCampaign campaign)
    {
        if(CampaignSender.Current()!=campaign.Sender)throw new InvalidOperationException("The sender settings changed. Prepare and review the campaign again.");
        StopDeliveryButton.Content="Stop after current email";
        bool live=campaign.Mode=="Live",reporting=true;
        int total=deliveryStore.Rows(campaign).Count(r=>r.State is "Not attempted" or "Failed"),accepted=0;
        void ShowProgress(string message,int done,int count)
        {
            StatusText.Text=message+" ("+done+" / "+count+")";
            DemoProgress.Value=count==0?0:100.0*done/count;
            if(!live)return;
            accepted=done;
            LiveDeliveryProgress.Value=count==0?0:100.0*done/count;
            LiveProgressCount.Text=$"Mail server accepted: {done} / {count} | Remaining: {Math.Max(0,count-done)}";
            LiveProgressStatus.Text=message;
        }
        if(live){LiveProgressPanel.Visibility=Visibility.Visible;ShowProgress("Connecting to the mail server",0,total);}
        var progress=new Progress<(string Message,int Done,int Total)>(p=>{if(reporting)ShowProgress(p.Message,p.Done,p.Total);});
        var before=deliveryStore.Rows(campaign).Count(r=>r.State=="Accepted");
        try
        {
            await new CampaignDeliveryService(deliveryStore,deliveryTransport).Send(campaign,(m,d,t)=>((IProgress<(string Message,int Done,int Total)>)progress).Report((m,d,t)),deliveryStop!.Token);
        }
        finally
        {
            // Ignore any queued progress callback after writing the final outcome.
            reporting=false;
            var rows=deliveryStore.Rows(campaign);
            string summary=$"{campaign.Mode}: {rows.Count(r=>r.State=="Accepted")} accepted by mail server, {rows.Count(r=>r.State=="Failed")} failed, {rows.Count(r=>r.State=="Uncertain")} uncertain, {rows.Count(r=>r.State=="Not attempted")} not attempted. Final delivery is not confirmed. See View history.";
            StatusText.Text=summary;
            if(live)
            {
                accepted=Math.Max(0,rows.Count(r=>r.State=="Accepted")-before);
                LiveDeliveryProgress.Value=total==0?0:100.0*accepted/total;
                LiveProgressCount.Text=$"Mail server accepted: {accepted} / {total} | Remaining: {Math.Max(0,total-accepted)}";
                LiveProgressStatus.Text=(accepted==total&&total>0?"Sending finished. ":"Sending stopped. ")+summary;
            }
        }
    }
    private async void SendLive_Click(object sender,RoutedEventArgs e)
    {
        if(deliveryBusy||preparedLive is null)return;
        try
        {
            if(CampaignPreparation.Fingerprint(ReadDraft(),CampaignSender.Current())!=preparedLive.Fingerprint)throw new InvalidOperationException("Recipients, content or source attachments changed. Prepare all messages again.");
            if(!ConfirmDelivery(preparedLive))return;await SendCampaign(preparedLive,true);
        }
        catch(Exception ex){Error(ex);}finally{if(deliveryBusy)DeliveryBusy(false);}
    }
    private async Task SendCampaign(PreparedCampaign campaign,bool currentDraft)
    {
        try
        {
            if(currentDraft)SentDraftStore.Register(campaign,draftPath,ReadDraft());
            DeliveryBusy(true);await Deliver(campaign);
        }
        finally{if(deliveryBusy)DeliveryBusy(false);}
        FinishSentDraft(campaign);
    }
    private void FinishSentDraft(PreparedCampaign campaign)
    {
        try
        {
            var result=SentDraftStore.Complete(campaign,deliveryStore);if(result is null)return;
            if(result.Changed){LiveProgressStatus.Text+=" The saved draft changed after sending began and has been kept.";return;}
            bool same=draftActive&&SentDraftStore.SameContent(ReadDraft(),result.Link.Snapshot);
            if(!same)return; // A different or subsequently edited draft stays open.
            bool history=DeliveryHistoryPanel.Visibility==Visibility.Visible;
            StartScreen();
            SentCompletionNotice.Text=$"All {campaign.Messages.Count} live message(s) were accepted by the mail server. This draft has left Open draft. View history contains the sending result.";
            SentCompletionNotice.Visibility=Visibility.Visible;StatusText.Text=SentCompletionNotice.Text;
            if(history){DeliveryHistory_Click(this,new RoutedEventArgs());DeliveryHistoryGrid.SelectedItem=DeliveryHistoryGrid.Items.Cast<CampaignRun>().FirstOrDefault(r=>r.Id==campaign.Id);}
        }
        catch(Exception ex){StatusText.Text="Sending finished, but the saved draft could not be cleared: "+ex.Message;LiveProgressStatus.Text+=" "+StatusText.Text;}
    }
    private void StopDelivery_Click(object sender,RoutedEventArgs e)=>Cancel();
    private void ShowPreparedAttachment()
    {
        if(preparedAttachments.TryGetValue(currentPreviewIndex,out var message))
        {
            PreviewAttachments.ItemsSource=message.Attachments.Count==0?new object[]{"No attachments"}:message.Attachments.Cast<object>().ToList();
            PreviewSubject.Text=message.Subject;PreviewBody.Text=message.Body;ShowHtmlPreview(message.HtmlBody,message.Body);
            PreviewTo.Text="To: "+message.To+(message.To!=message.OriginalTo?"\nTEST - Intended resident: "+message.OriginalTo:"");
        }
    }
    private void OpenPreparedAttachment_Click(object sender,RoutedEventArgs e)
    {
        if((sender as Button)?.CommandParameter is not CampaignAttachment file){StatusText.Text="Use Documents & attachments > Save generated attachments to inspect exported files in your chosen folder. Sending preparation creates its own attachment previews.";return;}
        try
        {
            if(!File.Exists(file.Path)||CampaignHash.File(file.Path)!=file.Hash)throw new IOException("The prepared file changed or is missing. Prepare the campaign again.");
            if(!new[]{".pdf",".docx",".xlsx",".txt",".png",".jpg",".jpeg"}.Contains(Path.GetExtension(file.Path).ToLowerInvariant()))throw new InvalidOperationException("This file type cannot be opened from the preview. Inspect its source using File Explorer.");
            Process.Start(new ProcessStartInfo(file.Path){UseShellExecute=true});
        }
        catch(Exception ex){Error(ex);}
    }
    private void DeliveryHistory_Click(object sender,RoutedEventArgs e)
    {
        if(deliveryBusy)return;bool show=DeliveryHistoryPanel.Visibility!=Visibility.Visible;
        if(show){deliveryStore.RecoverInterrupted();RefreshDeliveryHistory();}
        DeliveryHistoryPanel.Visibility=show?Visibility.Visible:Visibility.Collapsed;Steps.Visibility=show?Visibility.Collapsed:Visibility.Visible;
        HistoryNavigationButton.Content=show?"Back to feature":"View history";UpdateSteps();
    }
    private void RefreshDeliveryHistory()
    {
        string? runId=(DeliveryHistoryGrid.SelectedItem as CampaignRun)?.Id;
        string? messageId=(DeliveryMessagesGrid.SelectedItem as DeliveryRow)?.Message.Id;
        var runs=deliveryStore.Runs().Take(20).ToList();
        DeliveryHistoryGrid.ItemsSource=runs;
        DeliveryHistoryGrid.SelectedItem=runs.FirstOrDefault(r=>r.Id==runId)??runs.FirstOrDefault();
        if(DeliveryHistoryGrid.SelectedItem is null)DeliveryMessagesGrid.ItemsSource=null;
        if(messageId is not null&&DeliveryMessagesGrid.Items.Cast<DeliveryRow>().FirstOrDefault(r=>r.Message.Id==messageId) is {} selected)DeliveryMessagesGrid.SelectedItem=selected;
        UpdateDeliverySelection();
    }
    private void RefreshDeliveryHistory_Click(object sender,RoutedEventArgs e)=>RefreshDeliveryHistory();
    private void DeliveryRun_Changed(object sender,SelectionChangedEventArgs e)
    {
        DeliveryMessagesGrid.ItemsSource=DeliveryHistoryGrid.SelectedItem is CampaignRun run?deliveryStore.Rows(run.Campaign):null;
        if(DeliveryMessagesGrid.Items.Count>0)DeliveryMessagesGrid.SelectedIndex=0;
        UpdateDeliverySelection();
    }
    private void DeliveryMessage_Changed(object sender,SelectionChangedEventArgs e)=>UpdateDeliverySelection();
    private void UpdateDeliverySelection()
    {
        if(DeliveryDetailsButton is null||DeliverySelectionHint is null)return;
        var row=DeliveryMessagesGrid.SelectedItem as DeliveryRow;
        DeliveryDetailsButton.IsEnabled=row is not null;
        DeliverySelectionHint.Text=row is null?"Select a campaign above, then a message below to view its details.":"Selected message: unit "+row.Unit+" - "+row.To;
    }
    private void DeliveryDetails_Click(object sender,RoutedEventArgs e)
    {
        if(DeliveryMessagesGrid.SelectedItem is not DeliveryRow row){UpdateDeliverySelection();return;}
        new Window{Owner=OwnerWindow,Title="Email message details",Width=850,Height=650,WindowStartupLocation=WindowStartupLocation.CenterOwner,Content=new TextBox{IsReadOnly=true,TextWrapping=TextWrapping.Wrap,VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Padding=new Thickness(24),Text="Unit: "+row.Unit+"\nTo: "+row.To+"\nOriginal resident: "+row.Message.OriginalTo+"\nState: "+row.State+"\n"+row.Detail+"\nSubject: "+row.Subject+"\n\n"+row.Message.Body+"\n\nAttachments:\n"+string.Join("\n",row.Message.Attachments.Select(a=>a.Name))}}.ShowDialog();
    }
    private async void RetryDelivery_Click(object sender,RoutedEventArgs e)
    {
        if(deliveryBusy||DeliveryHistoryGrid.SelectedItem is not CampaignRun run||!ConfirmDelivery(run.Campaign))return;
        try{await SendCampaign(run.Campaign,false);}catch(Exception ex){Error(ex);}finally{if(deliveryBusy)DeliveryBusy(false);RefreshDeliveryHistory();}
    }
}
