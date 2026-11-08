using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
namespace PropertyManager.App.PersonalizedEmail;
internal sealed record SentDraftLink(string? DraftPath,string? SavedHash,CampaignDraft Snapshot);
internal sealed record SentDraftResult(SentDraftLink Link,bool Removed,bool Changed);
internal static class SentDraftStore
{
    private static readonly byte[] entropy=Encoding.UTF8.GetBytes("PropertyManagerDemo.SentDraft.v1");
    private static string Root=>Path.Combine(Services.AppSettings.DataDirectory,"PersonalizedEmail","SentDrafts");
    private static string FilePath(string id,bool pending)
    {
        if(!Guid.TryParseExact(id,"N",out _))throw new InvalidDataException("Invalid sending history identifier.");
        return Path.Combine(Root,id+(pending?".pending":".sent"));
    }
    internal static void Register(PreparedCampaign campaign,string? draftPath,CampaignDraft snapshot)
    {
        if(campaign.Mode!="Live")return;
        string? hash=draftPath is not null&&File.Exists(draftPath)?CampaignHash.File(draftPath):null;
        var link=new SentDraftLink(draftPath,hash,snapshot);
        Directory.CreateDirectory(Root);string path=FilePath(campaign.Id,true),temp=path+"."+Guid.NewGuid().ToString("N")+".tmp";
        try{File.WriteAllBytes(temp,ProtectedData.Protect(JsonSerializer.SerializeToUtf8Bytes(link),entropy,DataProtectionScope.CurrentUser));File.Move(temp,path,true);}
        finally{if(File.Exists(temp))File.Delete(temp);}
    }
    internal static SentDraftResult? Complete(PreparedCampaign campaign,CampaignDeliveryStore store)
    {
        if(campaign.Mode!="Live"||campaign.Messages.Count==0||store.Rows(campaign).Any(r=>r.State!="Accepted"))return null;
        string pending=FilePath(campaign.Id,true),sent=FilePath(campaign.Id,false);
        if(!File.Exists(pending)&&!File.Exists(sent))return null;
        var link=JsonSerializer.Deserialize<SentDraftLink>(ProtectedData.Unprotect(File.ReadAllBytes(File.Exists(pending)?pending:sent),entropy,DataProtectionScope.CurrentUser))??throw new InvalidDataException("Sent draft recovery data is empty.");
        // Keep a protected snapshot of the actual sent draft before removing its
        // saved version. Never remove a saved file that changed during delivery.
        if(File.Exists(pending))File.Move(pending,sent,true);
        bool changed=false,removed=false;
        if(link.DraftPath is not null&&File.Exists(link.DraftPath))
        {
            changed=link.SavedHash is null||CampaignHash.File(link.DraftPath)!=link.SavedHash;
            if(!changed){CampaignDraftLibrary.DeletePath(link.DraftPath);removed=true;}
        }
        return new(link,removed,changed);
    }
    internal static bool SameContent(CampaignDraft a,CampaignDraft b)
    {
        CampaignDraft Copy(CampaignDraft d){var copy=JsonSerializer.Deserialize<CampaignDraft>(JsonSerializer.Serialize(d))!;copy.CurrentStep=copy.UnlockedStep=null;return copy;}
        return JsonSerializer.Serialize(Copy(a))==JsonSerializer.Serialize(Copy(b));
    }
}
