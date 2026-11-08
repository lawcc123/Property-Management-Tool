using System.Security.Cryptography;
using System.Text.Json;
using PropertyManager.App.Operations;
namespace PropertyManager.App.PersonalizedEmail;
internal sealed record CampaignSender(string Server,int Port,string From)
{
    internal static CampaignSender Current(){var s=OperationsSettings.Load();s.ValidateSender();if(s.SmtpPort==465)throw new InvalidOperationException("Use a STARTTLS SMTP port, such as 587. SSL-only port 465 is not supported.");return new(s.SmtpServer,s.SmtpPort,s.From);}
}
internal sealed record CampaignAttachment(string Path,string Name,string Hash){public override string ToString()=>Name;}
internal sealed record DeliveryMessage(string Id,string Unit,string OriginalTo,string To,string Subject,string Body,List<CampaignAttachment> Attachments,string HtmlBody="");
internal sealed record PreparedCampaign(string Id,string Fingerprint,string Title,string Mode,string TestInbox,CampaignSender Sender,List<DeliveryMessage> Messages,List<string>? GeneratedCacheKeys=null);
internal sealed record CampaignRun(string Id,string Created,string Status,PreparedCampaign Campaign)
{
    public string LocalTime=>TimeZoneInfo.ConvertTime(DateTimeOffset.Parse(Created),TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")).ToString("yyyy-MM-dd HH:mm:ss");
    public string Title=>Campaign.Title;
    public string Mode=>Campaign.Mode;
    public int Messages=>Campaign.Messages.Count;
}
internal sealed record DeliveryRow(DeliveryMessage Message,string State,string Detail)
{
    public string Unit=>Message.Unit;public string To=>Message.To;public string Subject=>Message.Subject;
}
internal static class CampaignHash
{
    internal static string File(string path){using var stream=System.IO.File.OpenRead(path);return Convert.ToHexString(SHA256.HashData(stream));}
    internal static string Text(string text)=>Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(text)));
}
internal interface ICampaignDocuments:IDisposable
{
    string Generate(string template,string output,bool personalize,IReadOnlyDictionary<string,string> fields);
}
internal sealed class CampaignNotSubmittedException(string message):Exception(message);
internal interface ICampaignTransport:IAsyncDisposable
{
    Task Connect(PreparedCampaign campaign);
    Task Send(DeliveryMessage message,Action beforeSubmission);
}
