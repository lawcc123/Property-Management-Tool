using System.IO;
using System.Text;
namespace PropertyManager.App.PersonalizedEmail;

internal static class CampaignMessageSize
{
    // Reserve 5 MB below the common 25 MB SMTP ceiling, including MIME/base64 expansion.
    internal const long Limit=20_000_000;
    internal static string Display(long bytes)=>$"{bytes/1_000_000.0:0.00} MB";
    internal static long Estimate(IEnumerable<long> sizes,string subject,string body,string names="")
    {
        checked
        {
            long total=8192+Encoding.UTF8.GetByteCount(subject+body+names)*3L;
            foreach(long size in sizes){long encoded=((size+2)/3)*4;total+=encoded+((encoded+75)/76)*2+4096;}
            return total;
        }
    }
    internal static long Sources(CampaignDraft draft)
    {
        var sizes=draft.Documents.Select(d=>new FileInfo(d.Path).Length).ToArray();
        var ready=draft.Recipients.Where(r=>r.Included&&r.Status=="Ready").ToList();
        string names=string.Join("",draft.Documents.Select(d=>d.OutputName+d.FileName));
        string signature=draft.SignatureHtml.Length>0?draft.SignatureHtml:SignatureHtml.FromPlain(draft.Signature);
        if(ready.Count==0)return Estimate(sizes,draft.Subject,draft.Body+draft.Signature+SignatureHtml.Message(draft.Body,signature),names);
        return ready.Max(r=>Estimate(sizes,CampaignService.Render(draft.Subject,r.Values()),CampaignService.Render(draft.Body+"\n\n"+draft.Signature,r.Values())+SignatureHtml.Message(CampaignService.Render(draft.Body,r.Values()),SignatureHtml.Render(signature,r.Values())),names));
    }
    internal static void Validate(long size)
    {
        if(size>Limit)throw new InvalidDataException($"This message needs approximately {Display(size)}, exceeding the 20 MB message budget (including attachment encoding). Remove or reduce attachments before continuing.");
    }
    internal static void Validate(DeliveryMessage message)=>Validate(Estimate(message.Attachments.Select(a=>new FileInfo(a.Path).Length),message.Subject,message.Body+message.HtmlBody,string.Join("",message.Attachments.Select(a=>a.Name))+message.To));
}
