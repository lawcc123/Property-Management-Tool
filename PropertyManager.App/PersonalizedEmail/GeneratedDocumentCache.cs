using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
namespace PropertyManager.App.PersonalizedEmail;

// Disposable acceleration only: delivery snapshots and history never depend on it.
internal sealed class GeneratedDocumentCache
{
    internal const long MaxEntryBytes=24_000_000,MaxTotalBytes=128_000_000;
    internal string Root {get;}
    internal GeneratedDocumentCache(string? root=null)
    {
        Root=root??Path.Combine(Services.AppSettings.DataDirectory,"PersonalizedEmail","DocumentCache");
        Prune();
    }
    internal static string Key(string hash,bool personalize,string format,IReadOnlyDictionary<string,string> values)
        =>CampaignHash.Text(JsonSerializer.Serialize(new{Engine="word-documents-v2",Source=hash,Personalize=personalize,Format=format,
            Values=personalize?values.OrderBy(v=>v.Key,StringComparer.OrdinalIgnoreCase).Select(v=>new{Key=v.Key.ToUpperInvariant(),v.Value}).ToArray():[]}));
    private string PathFor(string key)=>Path.Combine(Root,key+".bgcache");
    private static byte[] Entropy(string key)=>Encoding.UTF8.GetBytes("PropertyManagerDemo.DocumentCache.v1:"+key);
    internal bool TryCopy(string key,string output)
    {
        byte[]? payload=null;
        try
        {
            string path=PathFor(key);var info=new FileInfo(path);
            if(!info.Exists||info.Length>MaxEntryBytes+8192||info.LastWriteTimeUtc<DateTime.UtcNow.AddDays(-7))return false;
            using(DocumentPerformance.Measure("cache.read"))payload=ProtectedData.Unprotect(File.ReadAllBytes(path),Entropy(key),DataProtectionScope.CurrentUser);
            if(payload.Length<=32||payload.Length>MaxEntryBytes+32||!CryptographicOperations.FixedTimeEquals(payload.AsSpan(0,32),SHA256.HashData(payload.AsSpan(32))))return false;
            using(var file=new FileStream(output,FileMode.CreateNew,FileAccess.Write,FileShare.None))file.Write(payload.AsSpan(32));
            try{File.SetLastWriteTimeUtc(path,DateTime.UtcNow);}catch(IOException){}catch(UnauthorizedAccessException){}
            DocumentPerformance.Record("cache.hit",bytes:payload.Length-32);return true;
        }
        catch(Exception ex) when(ex is IOException or UnauthorizedAccessException or CryptographicException){return false;}
        finally{if(payload is not null)CryptographicOperations.ZeroMemory(payload);}
    }
    internal void Store(string key,string source)
    {
        byte[]? data=null,payload=null;string? stage=null;
        try
        {
            if(new FileInfo(source).Length>MaxEntryBytes)return;
            using var timing=DocumentPerformance.Measure("cache.write");
            data=File.ReadAllBytes(source);payload=new byte[32+data.Length];SHA256.HashData(data).CopyTo(payload,0);data.CopyTo(payload,32);
            Directory.CreateDirectory(Root);stage=Path.Combine(Root,key+"."+Guid.NewGuid().ToString("N")+".tmp");
            File.WriteAllBytes(stage,ProtectedData.Protect(payload,Entropy(key),DataProtectionScope.CurrentUser));File.Move(stage,PathFor(key),true);stage=null;
            Prune();
        }
        catch(Exception ex) when(ex is IOException or UnauthorizedAccessException or CryptographicException){} // Cache is optional.
        finally
        {
            if(data is not null)CryptographicOperations.ZeroMemory(data);if(payload is not null)CryptographicOperations.ZeroMemory(payload);
            if(stage is not null)try{File.Delete(stage);}catch(IOException){}catch(UnauthorizedAccessException){}
        }
    }
    internal bool Remove(IEnumerable<string> keys)
    {
        bool complete=true;
        if(Directory.Exists(Root)&&(File.GetAttributes(Root)&FileAttributes.ReparsePoint)!=0)return false;
        foreach(string key in keys.Distinct(StringComparer.Ordinal))
        {
            if(key.Length!=64||key.Any(c=>!Uri.IsHexDigit(c))){complete=false;continue;}
            try{File.Delete(PathFor(key));}catch(IOException){complete=false;}catch(UnauthorizedAccessException){complete=false;}
        }
        return complete;
    }
    internal void Prune()
    {
        try
        {
            if(!Directory.Exists(Root))return;
            var files=new DirectoryInfo(Root).GetFiles("*.bgcache").OrderByDescending(f=>f.LastWriteTimeUtc).ToList();long size=0;
            foreach(var file in files)
            {
                size+=file.Length;
                if(file.LastWriteTimeUtc<DateTime.UtcNow.AddDays(-7)||size>MaxTotalBytes){size-=file.Length;file.Delete();}
            }
            foreach(var file in new DirectoryInfo(Root).GetFiles("*.tmp").Where(f=>f.LastWriteTimeUtc<DateTime.UtcNow.AddDays(-1)))file.Delete();
        }
        catch(IOException){}catch(UnauthorizedAccessException){}
    }
}
