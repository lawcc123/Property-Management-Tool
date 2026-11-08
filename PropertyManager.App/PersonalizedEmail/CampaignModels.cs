using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using PropertyManager.App.EmailLists;

namespace PropertyManager.App.PersonalizedEmail;

internal sealed class CampaignRecipient : INotifyPropertyChanged
{
    public string RecipientId {get;set;}=Guid.NewGuid().ToString("N");
    public Dictionary<string,string> Fields {get;set;}=new(StringComparer.OrdinalIgnoreCase);
    private string unit="",names="",email="",corporation="";private bool included=true;
    public string Unit {get=>unit;set{unit=value;Changed();}}
    public string Names {get=>names;set{names=value;Changed();}}
    public string Email {get=>email;set{email=value;Changed();}}
    public string Corporation {get=>corporation;set{corporation=value;Changed();}}
    public bool Included {get=>included;set{included=value;Changed();}}
    public string Status=>!Included?"Excluded":string.IsNullOrWhiteSpace(Email)?"Print only":GeneratedEmailListService.EmailParts(Email) is {Length:>0} parts&&parts.All(GeneratedEmailListService.IsValidEmail)?"Ready":"Check email";
    public event PropertyChangedEventHandler? PropertyChanged;
    private void Changed([System.Runtime.CompilerServices.CallerMemberName]string? property=null){PropertyChanged?.Invoke(this,new(property));PropertyChanged?.Invoke(this,new(nameof(Status)));}
    internal Dictionary<string,string> Values()=>new(Fields,StringComparer.OrdinalIgnoreCase){["Unit"]=Unit,["Names"]=Names,["Email"]=Email,["Corporation"]=Corporation};
}
internal sealed class CampaignDocument : INotifyPropertyChanged
{
    public string Path {get;set;}="";
    public string FileName=>System.IO.Path.GetFileName(Path);
    [System.Text.Json.Serialization.JsonIgnore] public string SizeText{get{try{return CampaignMessageSize.Display(new FileInfo(Path).Length);}catch{return "Unavailable";}}}
    internal void RefreshSize()=>PropertyChanged?.Invoke(this,new(nameof(SizeText)));
    public bool IsWord=>System.IO.Path.GetExtension(Path).Equals(".docx",StringComparison.OrdinalIgnoreCase);
    private bool personalize=true;private string format="PDF";
    public bool Personalize
    {
        get=>personalize;
        set
        {
            if(personalize==value)return;
            personalize=value;
            if(!value)OutputName=SharedOutputName(outputName);
            PropertyChanged?.Invoke(this,new(nameof(Personalize)));
            PropertyChanged?.Invoke(this,new(nameof(RequiresGeneration)));
            PropertyChanged?.Invoke(this,new(nameof(Treatment)));
        }
    }
    public string Format {get=>format;set{format=value;PropertyChanged?.Invoke(this,new(nameof(Format)));PropertyChanged?.Invoke(this,new(nameof(RequiresGeneration)));}}
    public bool RequiresGeneration=>IsWord&&(Personalize||Format=="PDF");
    public IEnumerable<string> OutputFormats=>IsWord?new[]{"PDF","Word"}:new[]{"Original"};
    public event PropertyChangedEventHandler? PropertyChanged;
    private string outputName="{{Unit}} - Notice";
    public string OutputName {get=>outputName;set{outputName=personalize?value:SharedOutputName(value);PropertyChanged?.Invoke(this,new(nameof(OutputName)));}}
    private string SharedOutputName(string value)
    {
        const string tokens=@"\{\{[^{}]+\}\}";
        if(!Regex.IsMatch(value,tokens))return value;
        string Clean(string name)=>Regex.Replace(Regex.Replace(name,tokens,""),@"\s+"," ").Trim(' ','-','\u2013','\u2014','_','|','.');
        string name=Clean(value);
        name=Regex.Replace(name,@"(?:\s*[-\u2013\u2014_|]\s*){2,}"," - ");
        if(name.Length==0)name=Clean(System.IO.Path.GetFileNameWithoutExtension(Path));
        return name.Length==0?"Notice":name;
    }
    public string Treatment=>IsWord?(Personalize?"Personalized Word template":"Shared Word document"):"Same file for everyone";
}
internal sealed record CampaignPreview(string Unit,string Names,string OriginalTo,string To,string Subject,string Body,List<string> Attachments,Dictionary<string,string> Fields,string HtmlBody="")
{
    public string Label=>Unit+" - "+Names;
}
internal sealed class CampaignDraft
{
    public int Version {get;set;}=1;
    public string Title {get;set;}="Untitled notice";
    public string LastExportFolder {get;set;}="";
    public int? CurrentStep {get;set;}
    public int? UnlockedStep {get;set;}
    public string Source {get;set;}="";
    public string Subject {get;set;}="Notice for {{Unit}}";
    public string Body {get;set;}="Dear {{Names}},\n\nWe are writing to you regarding unit {{Unit}}.\n\nThank you.";
    public string Signature {get;set;}="";
    public string SignatureHtml {get;set;}="";
    public bool TestMode {get;set;}
    public string TestInbox {get;set;}="";
    public List<CampaignRecipient> Recipients {get;set;}=[];
    public List<CampaignDocument> Documents {get;set;}=[];
}
internal static class CampaignService
{
    private static readonly Regex Tokens=new(@"\{\{(?<field>[^{}]+)\}\}",RegexOptions.Compiled);
    internal static List<string> Fields(string text)=>Tokens.Matches(text).Select(m=>m.Groups["field"].Value.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    internal static string Render(string text,Dictionary<string,string> values)=>Tokens.Replace(text,m=>values.TryGetValue(m.Groups["field"].Value.Trim(),out var value)?value:throw new InvalidDataException("Unknown personalization field: "+m.Groups["field"].Value));
    internal static List<CampaignRecipient> Map(RecipientFileData data,string emailColumn,string unitColumn,string nameColumn,string corporation="")
    {
        if(!data.Columns.Contains(emailColumn))throw new InvalidDataException("Choose the column containing email addresses.");
        string V(Dictionary<string,string> row,string key)=>row.FirstOrDefault(p=>p.Key.Equals(key,StringComparison.OrdinalIgnoreCase)).Value??"";
        return data.Rows.Select(row=>new CampaignRecipient{Fields=new(row,StringComparer.OrdinalIgnoreCase),Unit=V(row,unitColumn),Names=nameColumn=="(First + last name)"?string.Join(" ",new[]{V(row,"First Name"),V(row,"Last Name")}.Where(v=>v.Length>0)):V(row,nameColumn),Email=V(row,emailColumn),Corporation=corporation.Length>0?corporation:V(row,"Corporation") is {Length:>0} c?c:V(row,"Building Name")}).ToList();
    }
    internal static string DocumentText(string path,bool mergeFieldTokens=false)
    {
        using var archive=ZipFile.OpenRead(path);XNamespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var parts=archive.Entries.Where(e=>e.FullName.StartsWith("word/",StringComparison.Ordinal)&&e.FullName.EndsWith(".xml",StringComparison.OrdinalIgnoreCase)&&!e.FullName.Contains("/_rels/")).ToList();
        if(!parts.Any(e=>e.FullName=="word/document.xml"))throw new InvalidDataException("This is not a valid Word .docx document.");
        var text=new StringBuilder();
        foreach(var part in parts){using var stream=part.Open();var xml=XDocument.Load(stream);if(mergeFieldTokens)text.Append(WordPreviewText.Read(xml));else foreach(var p in xml.Descendants(w+"p"))text.AppendLine(string.Concat(p.Descendants(w+"t").Select(t=>t.Value)));}
        return text.ToString();
    }
    internal static List<string> DocumentFields(string path)
    {
        var fields=new List<string>();var text=new StringBuilder();using var archive=ZipFile.OpenRead(path);XNamespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        foreach(var part in archive.Entries.Where(e=>e.FullName.StartsWith("word/")&&e.FullName.EndsWith(".xml")))
        {
            using var stream=part.Open();var xml=XDocument.Load(stream);foreach(var paragraph in xml.Descendants(w+"p"))text.AppendLine(string.Concat(paragraph.Descendants(w+"t").Select(t=>t.Value)));var codes=string.Join(" ",xml.Descendants(w+"instrText").Select(t=>t.Value).Concat(xml.Descendants(w+"fldSimple").Select(t=>(string?)t.Attribute(w+"instr")??"")));
            foreach(Match match in Regex.Matches(codes,"MERGEFIELD\\s+(?:\"(?<field>[^\"]+)\"|(?<field>[^\\s\\\\]+))",RegexOptions.IgnoreCase))fields.Add(match.Groups["field"].Value);
        }
        fields.InsertRange(0,Fields(text.ToString()));
        if(archive.GetEntry("word/document.xml") is null)throw new InvalidDataException("This is not a valid Word .docx document.");
        return fields.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }
    internal static List<CampaignPreview> Prepare(CampaignDraft draft)
    {
        if(string.IsNullOrWhiteSpace(draft.Subject)||draft.Subject.Contains('\r')||draft.Subject.Contains('\n'))throw new InvalidDataException("Enter a subject on one line.");
        if(string.IsNullOrWhiteSpace(draft.Body))throw new InvalidDataException("Write your email message first.");
        if(draft.TestMode&&!GeneratedEmailListService.IsValidEmail(draft.TestInbox.Trim()))throw new InvalidDataException("Enter one valid test inbox address.");
        var docs=new Dictionary<CampaignDocument,List<string>>();
        foreach(var doc in draft.Documents)
        {
            if(!File.Exists(doc.Path))throw new FileNotFoundException("Attachment not found: "+doc.FileName);
            if(doc.IsWord&&doc.Format is not ("PDF" or "Word"))throw new InvalidDataException("Choose PDF or Word for "+doc.FileName);
            if(doc.IsWord&&string.IsNullOrWhiteSpace(doc.OutputName))throw new InvalidDataException("Enter an output filename for "+doc.FileName);
            docs[doc]=doc.IsWord&&doc.Personalize?DocumentFields(doc.Path):[];
        }
        if(draft.Recipients.Any(r=>r.Included&&r.Status=="Check email"))throw new InvalidDataException("Correct or exclude rows marked Check email before reviewing.");
        if(draft.SignatureHtml.Length>0)SignatureHtml.Parse(draft.SignatureHtml);
        var result=new List<CampaignPreview>();var seen=new HashSet<string>();
        foreach(var recipient in draft.Recipients.Where(r=>r.Included&&r.Status=="Ready"))
        {
            var values=recipient.Values();var fields=Fields(draft.Subject+"\n"+draft.Body+"\n"+draft.Signature+"\n"+draft.SignatureHtml).Concat(docs.SelectMany(d=>d.Value)).Concat(draft.Documents.Where(d=>d.RequiresGeneration).SelectMany(d=>Fields(d.OutputName))).Distinct(StringComparer.OrdinalIgnoreCase);
            foreach(string field in fields)if(!values.ContainsKey(field))throw new InvalidDataException("Field '"+field+"' is not available for unit "+recipient.Unit+". Choose a matching field or edit the template.");
            string subject=Render(draft.Subject,values),body=Render(draft.Body+(draft.Signature.Length>0?"\n\n"+draft.Signature:""),values);
            string htmlBody=SignatureHtml.Message(Render(draft.Body,values),SignatureHtml.Render(draft.SignatureHtml.Length>0?draft.SignatureHtml:SignatureHtml.FromPlain(draft.Signature),values));
            if(subject.Contains('\r')||subject.Contains('\n'))throw new InvalidDataException("A subject field contains a line break for unit "+recipient.Unit+".");
            var names=draft.Documents.Select(doc=>doc.RequiresGeneration?SafeName(doc.Personalize?Render(doc.OutputName,values):doc.OutputName)+(doc.Format=="PDF"?".pdf":".docx"):doc.FileName).ToList();
            var attachments=draft.Documents.Select((doc,index)=>names[index]+(doc.RequiresGeneration?" (planned; not generated)":"")).ToList();
            if(names.Distinct(StringComparer.OrdinalIgnoreCase).Count()!=names.Count)throw new InvalidDataException("Two attachments have the same output filename for unit "+recipient.Unit+". Give each document its own name.");
            foreach(string email in GeneratedEmailListService.EmailParts(recipient.Email))
            {
                string key=JsonSerializer.Serialize(new{recipient.Corporation,recipient.Unit,Email=email.ToLowerInvariant(),Subject=subject,Body=body,HtmlBody=htmlBody,Attachments=attachments,Values=values.OrderBy(v=>v.Key,StringComparer.OrdinalIgnoreCase)});if(!seen.Add(key))continue;
                result.Add(new(recipient.Unit,recipient.Names,email,draft.TestMode?draft.TestInbox.Trim():email,(draft.TestMode?"[TEST - Unit "+recipient.Unit+"] ":"")+subject,body,attachments,values,htmlBody));
            }
        }
        if(result.Count==0)throw new InvalidDataException("No included recipients have a valid email address. Print-only records remain in your draft.");
        return result;
    }
    internal static string SafeName(string text){foreach(char c in System.IO.Path.GetInvalidFileNameChars())text=text.Replace(c,'_');return string.IsNullOrWhiteSpace(text)?"Notice":text.Trim();}
}
internal static class CampaignDraftStore
{
    internal static string DirectoryPath=>Path.Combine(Services.AppSettings.DataDirectory,"PersonalizedEmail","Drafts");
    private static readonly byte[] Entropy=Encoding.UTF8.GetBytes("PropertyManagerDemo.PersonalizedEmail.Draft.v1");
    internal static string Save(CampaignDraft draft,string? path=null)
    {
        Directory.CreateDirectory(DirectoryPath);
        string title=draft.Title;foreach(char c in Path.GetInvalidFileNameChars())title=title.Replace(c,'_');title=title.Trim();if(title.Length==0)title="Untitled notice";if(title.Length>60)title=title[..60];
        string date=TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow,TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")).ToString("yyyyMMdd-HHmmss");
        path??=Path.Combine(DirectoryPath,title+" - "+date+" - "+Guid.NewGuid().ToString("N")[..6]+".bgdraft");string stage=path+"."+Guid.NewGuid().ToString("N")+".tmp";
        try{File.WriteAllBytes(stage,ProtectedData.Protect(JsonSerializer.SerializeToUtf8Bytes(draft),Entropy,DataProtectionScope.CurrentUser));File.Move(stage,path,true);}finally{if(File.Exists(stage))File.Delete(stage);}return path;
    }
    internal static CampaignDraft Load(string path)
    {
        var draft=JsonSerializer.Deserialize<CampaignDraft>(ProtectedData.Unprotect(File.ReadAllBytes(path),Entropy,DataProtectionScope.CurrentUser))??throw new InvalidDataException("The saved draft is empty.");
        if(draft.Version!=1)throw new InvalidDataException("This draft version is not supported.");foreach(var row in draft.Recipients)row.Fields=new(row.Fields,StringComparer.OrdinalIgnoreCase);return draft;
    }
}
