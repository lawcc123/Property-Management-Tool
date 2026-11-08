namespace PropertyManager.App.PersonalizedEmail;
internal sealed class CampaignDeliveryService(CampaignDeliveryStore store,Func<ICampaignTransport> transport,TimeSpan? interval=null)
{
    internal async Task Send(PreparedCampaign campaign,Action<string,int,int> progress,CancellationToken stop)
    {
        using var held=store.Lock();
        var pending=store.Rows(campaign).Where(r=>r.State is "Not attempted" or "Failed").ToList();
        if(pending.Count==0)throw new InvalidOperationException("No messages are eligible to send. Accepted and uncertain outcomes are never resent automatically.");
        CampaignPreparation.Verify(campaign);
        if(campaign.Mode=="Test"&&(campaign.Messages.Count<1||campaign.Messages.Count>5))throw new InvalidOperationException("A test send is limited to 1 to 5 selected messages.");
        store.SetRun(campaign.Id,"Sending");int sent=0;bool connected=false;
        await using var smtp=transport();
        try
        {
            stop.ThrowIfCancellationRequested();progress("Connecting to the mail server",0,pending.Count);await smtp.Connect(campaign);connected=true;
            foreach(var row in pending)
            {
                if(stop.IsCancellationRequested){store.SetRun(campaign.Id,"Stopped - remaining messages not attempted");return;}
                bool submitting=false;
                try
                {
                    progress("Sending unit "+row.Unit,sent,pending.Count);
                    await smtp.Send(row.Message,()=>{store.SetMessage(row.Message.Id,"Submitting","Submission started; check outcome before retrying.");submitting=true;});
                    store.SetMessage(row.Message.Id,"Accepted","Accepted by the mail server; final delivery is not confirmed.");sent++;progress("Mail server accepted unit "+row.Unit,sent,pending.Count);
                }
                catch(Exception ex)
                {
                    store.SetMessage(row.Message.Id,submitting&&ex is not CampaignNotSubmittedException?"Uncertain":"Failed",ex.Message);
                    store.SetRun(campaign.Id,"Stopped - review failure/uncertain outcome");throw;
                }
                if(sent<pending.Count)
                {
                    try{await Task.Delay(interval??TimeSpan.FromSeconds(3),stop);}catch(OperationCanceledException){store.SetRun(campaign.Id,"Stopped - remaining messages not attempted");return;}
                }
            }
            var rows=store.Rows(campaign);store.SetRun(campaign.Id,rows.All(r=>r.State=="Accepted")?"Accepted by mail server":"Finished - review remaining outcomes");
            CompletedCampaignFiles.TryCleanup(campaign,store);
        }
        catch(OperationCanceledException){store.SetRun(campaign.Id,"Stopped - no further messages attempted");}
        catch{if(!connected)store.SetRun(campaign.Id,"Connection failed - no new messages submitted");throw;}
    }
}
