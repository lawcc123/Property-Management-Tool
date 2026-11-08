using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
namespace PropertyManager.App.PersonalizedEmail;

// One short local template snapshot per source and one Word instance per operation.
internal sealed class CampaignDocumentBatch(Func<ICampaignDocuments> factory,string? cacheRoot=null):IDisposable
{
    private sealed record Source(string Original,string Local,string Hash,bool Cacheable);
    private readonly Dictionary<string,Source> sources=new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string,string> generated=new(StringComparer.Ordinal);
    private readonly GeneratedDocumentCache cache=new(cacheRoot);
    private readonly string scratch=Path.Combine(Path.GetTempPath(),"BGdoc-"+Guid.NewGuid().ToString("N"));
    private ICampaignDocuments? documents;
    private readonly HashSet<string> cacheKeys=new(StringComparer.Ordinal);
    internal List<string> CacheKeys=>cacheKeys.ToList();
    internal string Snapshot(CampaignDocument doc,string? expectedHash=null)
    {
        string original=Path.GetFullPath(doc.Path);
        if(sources.TryGetValue(original,out var prior))return prior.Local;
        Directory.CreateDirectory(scratch);
        string hash=expectedHash??CampaignHash.File(original),local=Path.Combine(scratch,"source-"+sources.Count+Path.GetExtension(original));
        using(DocumentPerformance.Measure("source.snapshot"))File.Copy(original,local,false);
        if(CampaignHash.File(local)!=hash||CampaignHash.File(original)!=hash)throw new IOException("A source attachment changed while copying. Prepare again.");
        sources.Add(original,new(original,local,hash,doc.IsWord&&CanCache(local)));
        return local;
    }
    // Fields that Word can update from time, other documents or external sources
    // must be regenerated. Hyperlinks and page numbers are stable for this input.
    private static bool CanCache(string path)
    {
        try
        {
            using var zip=ZipFile.OpenRead(path);XNamespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            foreach(var part in zip.Entries.Where(e=>e.FullName.EndsWith(".xml",StringComparison.OrdinalIgnoreCase)||e.FullName.EndsWith(".rels",StringComparison.OrdinalIgnoreCase)))
            {
                using var stream=part.Open();var xml=XDocument.Load(stream);
                if(part.FullName.EndsWith(".rels",StringComparison.OrdinalIgnoreCase)&&xml.Descendants().Any(e=>(string?)e.Attribute("TargetMode")=="External"&&!((string?)e.Attribute("Type")??"").EndsWith("/hyperlink",StringComparison.Ordinal)))return false;
                var codes=xml.Descendants(w+"instrText").Select(t=>t.Value).Concat(xml.Descendants(w+"fldSimple").Select(t=>(string?)t.Attribute(w+"instr")??""));
                foreach(string code in codes)
                    if(!Regex.IsMatch(code.Trim(),@"^(MERGEFIELD|PAGE|NUMPAGES|SECTIONPAGES|HYPERLINK)\b",RegexOptions.IgnoreCase))return false;
            }
            return true;
        }
        catch(InvalidDataException){return false;}
    }
    internal bool Generate(CampaignDocument doc,IReadOnlyDictionary<string,string> values,string output,CancellationToken token)
    {
        token.ThrowIfCancellationRequested();Snapshot(doc);
        var source=sources[Path.GetFullPath(doc.Path)];string format=doc.Format=="PDF"?"PDF":"Word";
        string key=GeneratedDocumentCache.Key(source.Hash,doc.Personalize,format,values);
        if(source.Cacheable)cacheKeys.Add(key);
        if(!generated.TryGetValue(key,out string? local))
        {
            local=Path.Combine(scratch,"result-"+Guid.NewGuid().ToString("N")+(format=="PDF"?".pdf":".docx"));
            bool reused=source.Cacheable&&cache.TryCopy(key,local);
            if(!reused)
            {
                DocumentPerformance.Record("cache.miss");
                documents??=factory();documents.Generate(source.Local,local,doc.Personalize,values);
                token.ThrowIfCancellationRequested();
                if(!File.Exists(local)||new FileInfo(local).Length==0)throw new IOException("A personalized attachment was not generated.");
                if(source.Cacheable)cache.Store(key,local);
            }
            if(source.Cacheable)generated[key]=local;
            token.ThrowIfCancellationRequested();
            using(DocumentPerformance.Measure("output.copy"))File.Copy(local,output,false);
            return reused;
        }
        token.ThrowIfCancellationRequested();
        using(DocumentPerformance.Measure("output.copy"))File.Copy(local,output,false);
        DocumentPerformance.Record("cache.batch.hit");return true;
    }
    internal void VerifySources()
    {
        using var timing=DocumentPerformance.Measure("source.verify");
        foreach(var source in sources.Values)if(CampaignHash.File(source.Original)!=source.Hash)throw new IOException("A source attachment changed during generation. Prepare again.");
    }
    public void Dispose()
    {
        try{documents?.Dispose();}
        finally
        {
            // Only files in this newly-created, private operation directory.
            try{if(Directory.Exists(scratch)){foreach(string file in Directory.GetFiles(scratch))File.Delete(file);Directory.Delete(scratch,false);}}catch(IOException){}catch(UnauthorizedAccessException){}
        }
    }
}
