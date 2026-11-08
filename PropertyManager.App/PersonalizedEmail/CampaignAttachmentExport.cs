using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
namespace PropertyManager.App.PersonalizedEmail;
internal sealed record AttachmentExportResult(string Folder,int Recipients,int Files,bool Complete,string Problem="");
internal sealed class CampaignAttachmentExport(Func<ICampaignDocuments> generator,string? cacheRoot=null)
{
    internal static string FileName(CampaignDocument doc,Dictionary<string,string> fields)
        =>SafeName(doc.RequiresGeneration?(doc.Personalize?CampaignService.Render(doc.OutputName,fields):doc.OutputName)+(doc.Format=="PDF"?".pdf":".docx"):doc.FileName);
    static string SafeName(string value)
    {
        string name=CampaignService.SafeName(value).TrimEnd(' ','.');
        if(name.Length==0)name="Attachment";
        if(Regex.IsMatch(Path.GetFileNameWithoutExtension(name),@"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$",RegexOptions.IgnoreCase))name="_"+name;
        return name;
    }
    static string Part(string value)=>SafeName(value).Length>80?SafeName(value)[..80]:SafeName(value);
    internal AttachmentExportResult Generate(CampaignDraft draft,string destination,Action<string,int,int> progress,CancellationToken token)
    {
        using var operationTiming=DocumentPerformance.Measure("export.total");
        draft=JsonSerializer.Deserialize<CampaignDraft>(JsonSerializer.Serialize(draft))!;
        if(!Directory.Exists(destination))throw new DirectoryNotFoundException("Choose an available destination folder.");
        var templates=draft.Documents.Where(d=>d.IsWord&&d.Personalize).ToList();
        if(templates.Count==0)throw new InvalidOperationException("Add a Word template with Personalize checked first. Shared files are not exported.");
        // Export every included recipient, including print-only records. Identical
        // snapshots produce one set; distinct units/merge values stay separate.
        var recipients=draft.Recipients.Where(r=>r.Included).GroupBy(r=>JsonSerializer.Serialize(r.Values().OrderBy(p=>p.Key,StringComparer.OrdinalIgnoreCase))).Select(g=>g.First()).ToList();
        if(recipients.Count==0)throw new InvalidOperationException("Include at least one recipient first.");
        if(recipients.Count>10000)throw new InvalidOperationException("Export up to 10,000 recipient sets at a time.");
        var hashes=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        foreach(var doc in templates)
        {
            token.ThrowIfCancellationRequested();
            if(!File.Exists(doc.Path)||new FileInfo(doc.Path).Length==0)throw new IOException("The source attachment is missing or empty: "+doc.FileName);
            if(doc.IsWord&&doc.Format is not ("PDF" or "Word"))throw new InvalidDataException("Choose PDF or Word for "+doc.FileName);
            if(doc.RequiresGeneration&&string.IsNullOrWhiteSpace(doc.OutputName))throw new InvalidDataException("Enter an output filename for "+doc.FileName);
            hashes[doc.Path]=CampaignHash.File(doc.Path);
            var fields=doc.Personalize&&doc.IsWord?CampaignService.DocumentFields(doc.Path):[];
            foreach(var row in recipients)foreach(string field in fields)if(!row.Values().ContainsKey(field))throw new InvalidDataException("Missing template field '"+field+"' for unit "+row.Unit+".");
        }
        token.ThrowIfCancellationRequested();
        var now=TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow,TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time"));
        string name=Part(draft.Title)+" - "+now.ToString("yyyyMMdd-HHmmss")+" - "+Guid.NewGuid().ToString("N")[..10];
        string folder=Path.Combine(Path.GetFullPath(destination),name+" - INCOMPLETE"),completeFolder=Path.Combine(Path.GetFullPath(destination),name);
        if(Directory.Exists(folder)||Directory.Exists(completeFolder))throw new IOException("The export folder already exists. Choose Save generated attachments again.");
        Directory.CreateDirectory(folder);int files=0,sets=0;
        using var batch=new CampaignDocumentBatch(generator,cacheRoot);
        try
        {
            try
            {
                for(int templateIndex=0;templateIndex<templates.Count;templateIndex++)
                {
                    var doc=templates[templateIndex];token.ThrowIfCancellationRequested();
                    batch.Snapshot(doc,hashes[doc.Path]);
                    string templateFolder=Path.Combine(folder,(templateIndex+1).ToString("D2")+" - "+Part(Path.GetFileNameWithoutExtension(doc.FileName)));
                    Directory.CreateDirectory(templateFolder);
                    var usedNames=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach(var row in recipients)
                    {
                        token.ThrowIfCancellationRequested();
                        string unit=string.IsNullOrWhiteSpace(row.Unit)?"Recipient":row.Unit;
                        string fileName=FileName(doc,row.Values());
                        // Output names remain as configured when unique. Shared output
                        // names must never overwrite another unit's personalized file.
                        string stem=Path.GetFileNameWithoutExtension(fileName),extension=Path.GetExtension(fileName);
                        if(!usedNames.Add(fileName))
                        {
                            string suffix=Part(string.Join(" - ",new[]{row.Corporation,"Unit "+unit}.Where(v=>v.Length>0)));
                            fileName=SafeName(stem+" - "+suffix+extension);int duplicate=2;
                            while(!usedNames.Add(fileName))fileName=SafeName(stem+" - "+suffix+" - "+(duplicate++)+extension);
                        }
                        string output=Path.Combine(templateFolder,fileName);
                        progress("Saving "+fileName+" for unit "+unit,files,recipients.Count*templates.Count);
                        if(batch.Generate(doc,row.Values(),output,token))progress("Reused unchanged "+fileName,files,recipients.Count*templates.Count);
                        if(!File.Exists(output)||new FileInfo(output).Length==0)throw new IOException("An output file was not generated: "+fileName);
                        files++;progress("Saved "+fileName,files,recipients.Count*templates.Count);
                    }
                }
                sets=recipients.Count;
                foreach(var source in hashes)if(CampaignHash.File(source.Key)!=source.Value)throw new IOException("A source file changed during export. Export again.");
                token.ThrowIfCancellationRequested();
            }
            finally{batch.VerifySources();}
            File.WriteAllText(Path.Combine(folder,"Export details.txt"),$"Draft: {draft.Title}\nExported (Example City): {now:yyyy-MM-dd HH:mm:ss zzz}\nRecipient sets: {sets}\nFiles: {files}\nStatus: Complete\n\nThese files were saved for storage/review. No emails were sent. Each numbered template folder contains its personalized Word/PDF files for all included recipients. Shared files are excluded. You can copy or delete this export folder using File Explorer.\n");
            Directory.Move(folder,completeFolder);return new(completeFolder,sets,files,true);
        }
        catch(Exception ex)
        {
            string problem=ex is OperationCanceledException?"Export cancelled. Some files may have been generated; this folder is incomplete.":"Export failed. This folder is incomplete. "+ex.Message;
            try{File.WriteAllText(Path.Combine(folder,"INCOMPLETE - read me.txt"),problem+"\nFiles saved: "+files+"\nNo emails were sent. You may inspect or delete this export folder using File Explorer.");}catch{}
            return new(folder,sets,files,false,problem);
        }
    }
}
