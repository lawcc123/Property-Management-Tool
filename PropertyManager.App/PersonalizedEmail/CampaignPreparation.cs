using System.IO;
using System.Text.Json;
namespace PropertyManager.App.PersonalizedEmail;
internal sealed class CampaignPreparation(CampaignDeliveryStore store,Func<ICampaignDocuments> documents,string? cacheRoot=null)
{
    internal static string Fingerprint(CampaignDraft draft,CampaignSender sender)
    {
        var copy=JsonSerializer.Deserialize<CampaignDraft>(JsonSerializer.Serialize(draft))!;copy.TestMode=false;
        return Fingerprint(copy,sender,CampaignService.Prepare(copy));
    }
    private static string Fingerprint(CampaignDraft draft,CampaignSender sender,IReadOnlyList<CampaignPreview> previews)
        =>CampaignHash.Text(JsonSerializer.Serialize(new{sender,Messages=previews,Files=draft.Documents.Select(d=>new{d.Personalize,d.Format,d.OutputName,Hash=CampaignHash.File(d.Path)})}));
    internal PreparedCampaign Prepare(CampaignDraft draft,CampaignSender sender,IReadOnlyList<int>? selected,string? testInbox,Action<string,int,int> progress,CancellationToken token)
    {
        using var operationTiming=DocumentPerformance.Measure("preparation.total");
        using var held=store.Lock();
        draft=JsonSerializer.Deserialize<CampaignDraft>(JsonSerializer.Serialize(draft))!;draft.TestMode=false;
        CampaignMessageSize.Validate(CampaignMessageSize.Sources(draft));
        var previews=CampaignService.Prepare(draft);string fingerprint=Fingerprint(draft,sender,previews);bool test=selected is not null;
        if(test&&(selected!.Count<1||selected.Count>5||selected.Distinct().Count()!=selected.Count||selected.Any(i=>i<0||i>=previews.Count)||!EmailLists.GeneratedEmailListService.IsValidEmail(testInbox??"")))throw new InvalidOperationException("Select 1 to 5 messages and enter one valid test inbox address.");
        if(!test&&draft.Source.StartsWith("Sample data",StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("Sample data can be sent only as selected tests. Load real recipients before live sending.");
        if(!test&&store.ExistingLive(fingerprint) is {} existing){VerifyForDelivery(existing,store);return existing;}
        string id=Guid.NewGuid().ToString("N"),root=Path.Combine(store.DirectoryPath,"Files",id);Directory.CreateDirectory(root);
        using var batch=new CampaignDocumentBatch(documents,cacheRoot);
        var sources=draft.Documents.Select(d=>{token.ThrowIfCancellationRequested();return batch.Snapshot(d);}).ToList();
        var indices=test?selected!.ToList():Enumerable.Range(0,previews.Count).ToList();var messages=new List<DeliveryMessage>();
        try
        {
            foreach(int index in indices)
            {
                token.ThrowIfCancellationRequested();var preview=previews[index];progress("Preparing unit "+preview.Unit,messages.Count,indices.Count);
                string folder=Path.Combine(root,(index+1).ToString("D5"));Directory.CreateDirectory(folder);var attachments=new List<CampaignAttachment>();
                for(int d=0;d<draft.Documents.Count;d++)
                {
                    token.ThrowIfCancellationRequested();var doc=draft.Documents[d];string name=doc.RequiresGeneration?(doc.Personalize?CampaignService.Render(doc.OutputName,preview.Fields):doc.OutputName)+(doc.Format=="PDF"?".pdf":".docx"):doc.FileName;
                    name=CampaignService.SafeName(name);string output=Path.Combine(folder,name);
                    progress("Preparing "+name+" for unit "+preview.Unit,messages.Count,indices.Count);
                    if(doc.RequiresGeneration){if(batch.Generate(doc,preview.Fields,output,token))progress("Reused unchanged "+name+" for unit "+preview.Unit,messages.Count,indices.Count);}else File.Copy(sources[d],output,false);
                    if(!File.Exists(output)||new FileInfo(output).Length==0)throw new IOException("An attachment could not be generated: "+name);
                    attachments.Add(new(output,name,CampaignHash.File(output)));
                }
                string body=test?"TEST ONLY - Unit "+preview.Unit+"\nIntended resident address: "+preview.OriginalTo+"\nActual destination: "+testInbox+"\n\n"+preview.Body:preview.Body;
                string html=test?SignatureHtml.TestBanner(preview.HtmlBody,"TEST ONLY - Unit "+preview.Unit+"\nIntended resident address: "+preview.OriginalTo+"\nActual destination: "+testInbox):preview.HtmlBody;
                messages.Add(new(Guid.NewGuid().ToString("N"),preview.Unit,preview.OriginalTo,test?testInbox!:preview.OriginalTo,test?"[TEST - Unit "+preview.Unit+"] "+preview.Subject:preview.Subject,body,attachments,html));
                progress("Prepared unit "+preview.Unit,messages.Count,indices.Count);
            }
        }
        finally{batch.VerifySources();}
        if(Fingerprint(draft,sender,previews)!=fingerprint)throw new IOException("Source attachments changed while generating documents. Prepare again before sending.");
        var campaign=new PreparedCampaign(id,fingerprint,draft.Title,test?"Test":"Live",test?testInbox!:"",sender,messages,batch.CacheKeys);Verify(campaign);store.Add(campaign);return campaign;
    }
    internal static void VerifyForDelivery(PreparedCampaign campaign,CampaignDeliveryStore store)
    {
        if(campaign.Mode=="Live"&&store.Rows(campaign).All(r=>r.State=="Accepted"))return;
        Verify(campaign);
    }
    internal static void Verify(PreparedCampaign campaign)
    {
        foreach(var message in campaign.Messages)
        {
            CampaignMessageSize.Validate(message);
            if(!EmailLists.GeneratedEmailListService.IsValidEmail(message.To)||message.Subject.Contains('\r')||message.Subject.Contains('\n'))throw new InvalidDataException("A prepared recipient or subject is invalid.");
            if(campaign.Mode=="Test"&&(!message.To.Equals(campaign.TestInbox,StringComparison.OrdinalIgnoreCase)||!message.Subject.StartsWith("[TEST - Unit ")))throw new InvalidDataException("The test destination was not redirected safely.");
            foreach(var file in message.Attachments)if(!File.Exists(file.Path)||CampaignHash.File(file.Path)!=file.Hash)throw new IOException("A prepared attachment changed or is missing. Prepare the campaign again: "+file.Name);
        }
    }
}
