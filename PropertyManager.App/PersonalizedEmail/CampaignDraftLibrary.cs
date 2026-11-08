using System.IO;
using System.Security.Cryptography;
using System.Text.Json;
namespace PropertyManager.App.PersonalizedEmail;
internal sealed record CampaignDraftEntry(string Path,string Title,string Subject,DateTimeOffset Modified,int Recipients,int Documents,bool CanOpen,string Status)
{
    public string Availability=>CanOpen?"Ready":"Unavailable";
    public string SavedLocal=>TimeZoneInfo.ConvertTime(Modified,TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")).ToString("yyyy-MM-dd HH:mm");
}
internal static class CampaignDraftLibrary
{
    private static string ManagedPath(string path)
    {
        string full=System.IO.Path.GetFullPath(path);
        if(!string.Equals(System.IO.Path.GetDirectoryName(full),System.IO.Path.GetFullPath(CampaignDraftStore.DirectoryPath),StringComparison.OrdinalIgnoreCase)||!System.IO.Path.GetExtension(full).Equals(".bgdraft",StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Choose a saved draft from this app's draft list.");
        if(File.Exists(full)&&(File.GetAttributes(full)&FileAttributes.ReparsePoint)!=0)throw new InvalidOperationException("Linked draft files cannot be managed here.");
        return full;
    }
    internal static List<CampaignDraftEntry> List()
    {
        if(!Directory.Exists(CampaignDraftStore.DirectoryPath))return [];
        var entries=new List<CampaignDraftEntry>();
        foreach(string path in Directory.EnumerateFiles(CampaignDraftStore.DirectoryPath,"*.bgdraft",SearchOption.TopDirectoryOnly))
        {
            var modified=new DateTimeOffset(File.GetLastWriteTimeUtc(path));
            try
            {
                var draft=CampaignDraftStore.Load(ManagedPath(path));
                entries.Add(new(path,draft.Title,draft.Subject,modified,draft.Recipients.Count,draft.Documents.Count,true,"Ready"));
            }
            catch(Exception ex)when(ex is IOException or UnauthorizedAccessException or CryptographicException or JsonException or InvalidOperationException or ArgumentException)
            {entries.Add(new(path,System.IO.Path.GetFileNameWithoutExtension(path),"",modified,0,0,false,"Cannot open with this Windows account, or file is damaged"));}
        }
        return entries.OrderByDescending(e=>e.Modified).ThenBy(e=>e.Title,StringComparer.OrdinalIgnoreCase).ToList();
    }
    internal static CampaignDraft Load(CampaignDraftEntry entry)=>CampaignDraftStore.Load(ManagedPath(entry.Path));
    internal static void Delete(CampaignDraftEntry entry)=>DeletePath(entry.Path);
    internal static void DeletePath(string path)=>File.Delete(ManagedPath(path));
}
