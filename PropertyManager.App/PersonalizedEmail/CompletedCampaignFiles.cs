using System.IO;
namespace PropertyManager.App.PersonalizedEmail;
internal static class CompletedCampaignFiles
{
    // Call only while holding the sending-store lock. Cleanup failures never
    // change mail acceptance states; startup recovery retries pending cleanup.
    internal static bool TryCleanup(PreparedCampaign campaign,CampaignDeliveryStore store,string? cacheRoot=null)
    {
        try
        {
            if(campaign.Mode!="Live"||campaign.Messages.Count==0||store.Rows(campaign).Any(r=>r.State!="Accepted"))return false;
            if(store.FilesCleaned(campaign.Id))return true;
            if(!Guid.TryParseExact(campaign.Id,"N",out _))return false;
            string root=Path.GetFullPath(Path.Combine(store.DirectoryPath,"Files"));
            string target=Path.GetFullPath(Path.Combine(root,campaign.Id));
            if(!string.Equals(Path.GetDirectoryName(target),root,StringComparison.OrdinalIgnoreCase))return false;
            if(Directory.Exists(root)&&(File.GetAttributes(root)&FileAttributes.ReparsePoint)!=0)return false;
            if(Directory.Exists(target))
            {
                var files=new List<string>();var directories=new List<string>();
                void Inspect(string directory)
                {
                    if((File.GetAttributes(directory)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("Linked delivery folders cannot be cleaned automatically.");
                    directories.Add(directory);
                    foreach(string file in Directory.GetFiles(directory))
                    {
                        if((File.GetAttributes(file)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("Linked delivery files cannot be cleaned automatically.");
                        files.Add(file);
                    }
                    foreach(string child in Directory.GetDirectories(directory))Inspect(child);
                }
                Inspect(target); // Check the complete owned tree before deleting anything.
                foreach(string file in files)File.Delete(file);
                foreach(string directory in directories.AsEnumerable().Reverse())Directory.Delete(directory,false);
            }
            if(!new GeneratedDocumentCache(cacheRoot).Remove(campaign.GeneratedCacheKeys??[]))return false;
            store.MarkFilesCleaned(campaign);return true;
        }
        catch(Exception ex)when(ex is IOException or UnauthorizedAccessException or ArgumentException or Microsoft.Data.Sqlite.SqliteException)
        {DocumentPerformance.Record("delivery.cleanup.pending");return false;}
    }
}
