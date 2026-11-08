using System.IO;
using System.Text.Json;

namespace PropertyManager.App.Envelopes;

internal sealed class EnvelopeStore
{
    internal string Root {get;}
    private readonly string seeds;
    private static readonly JsonSerializerOptions Json=new(){WriteIndented=true};
    internal EnvelopeStore(string? root=null,string? seeds=null)
    {
        Root=root??Path.Combine(Services.AppSettings.DataDirectory,"EnvelopePrinter");
        this.seeds=seeds??Path.Combine(AppContext.BaseDirectory,"Assets","EnvelopeDefaults");
    }
    internal AddressBook LoadBook()
    {
        Directory.CreateDirectory(Root);string path=Path.Combine(Root,"address-book.json");
        if(Settings.SettingsDatabase.Read(path) is {} json)
        {var existing=JsonSerializer.Deserialize<AddressBook>(json)??throw new InvalidDataException("The address book is empty or invalid.");existing.Validate();return existing;}
        var contractors=JsonSerializer.Deserialize<List<MailingAddress>>(File.ReadAllText(Path.Combine(seeds,"contractors.json")))??throw new InvalidDataException("Contractor defaults are invalid.");
        using var settings=JsonDocument.Parse(File.ReadAllText(Path.Combine(seeds,"settings.json")));
        var offices=JsonSerializer.Deserialize<List<MailingAddress>>(settings.RootElement.GetProperty("offices").GetRawText())??[];
        var book=new AddressBook{Contractors=contractors,Offices=offices};book.Validate();
        Settings.SettingsDatabase.Write(path,JsonSerializer.Serialize(book,Json),onlyIfMissing:true);
        return JsonSerializer.Deserialize<AddressBook>(Settings.SettingsDatabase.Read(path)!)!;
    }
    internal void SaveBook(AddressBook book){book.Validate();Write(Path.Combine(Root,"address-book.json"),book);}
    internal EnvelopeOptions LoadOptions()=>(Read<EnvelopeOptions>(Path.Combine(Root,"last-options.json"))??new()).CurrentEnvelopeSettings();
    internal void SaveOptions(EnvelopeOptions options){options.Validate();Write(Path.Combine(Root,"last-options.json"),options);}
    internal void SaveProfile(EnvelopeOptions options)
    {
        options.Validate();if(string.IsNullOrWhiteSpace(options.Printer))throw new InvalidOperationException("Select a printer first.");
        var profiles=Read<Dictionary<string,EnvelopeOptions>>(Path.Combine(Root,"printer-profiles.json"))??new();
        profiles[options.Printer+"|"+options.Preset]=options.Copy();Write(Path.Combine(Root,"printer-profiles.json"),profiles);SaveOptions(options);
    }
    internal EnvelopeOptions? LoadProfile(string printer,string preset)
    {var profiles=Read<Dictionary<string,EnvelopeOptions>>(Path.Combine(Root,"printer-profiles.json"));return profiles?.GetValueOrDefault(printer+"|"+preset)?.CurrentEnvelopeSettings();}
    internal EnvelopeHistoryEntry Create(EnvelopeSnapshot snapshot,string kind,string? parent=null)
    {
        if(kind=="Preview")throw new InvalidOperationException("Preview history is no longer supported. Previewing does not create a history record.");
        snapshot.Validate();var entry=new EnvelopeHistoryEntry{Snapshot=snapshot.Copy(),Kind=kind,ParentId=parent};
        entry.DocumentName="PropertyManager envelopes "+entry.Id;
        entry.Events.Add(new(DateTimeOffset.UtcNow,"Requested",kind=="Word document"?"Word document requested. No print submission.":"Print requested; not yet submitted."));
        Write(HistoryPath(entry.Id),entry);return entry;
    }
    internal EnvelopeHistoryEntry ReadHistory(string id)
    {
        var entry=Read<EnvelopeHistoryEntry>(HistoryPath(id))??throw new FileNotFoundException("History entry not found.");
        ValidateEntry(entry,id);return entry;
    }
    private static void ValidateEntry(EnvelopeHistoryEntry entry,string id)
    {
        if(entry.Id!=id||entry.Events is null||entry.Events.Count==0||entry.Snapshot is null||entry.DocumentName!="PropertyManager envelopes "+id)throw new InvalidDataException("History entry is incomplete or inconsistent.");
        try{entry.Snapshot.Validate();}catch(Exception ex){throw new InvalidDataException("History snapshot is invalid.",ex);}
    }
    internal EnvelopeHistoryEntry Update(string id,string status,string details,int? jobId=null)
    {
        Directory.CreateDirectory(Root);
        using var held=LockHistory();
        var entry=ReadHistory(id);entry.JobId=jobId??entry.JobId;entry.Events.Add(new(DateTimeOffset.UtcNow,status,details));Write(HistoryPath(id),entry);return entry;
    }
    internal void DeleteHistory(string id)
    {
        string path=HistoryPath(id); // Validate the ID before touching the filesystem.
        Directory.CreateDirectory(Root);
        using var held=LockHistory();
        File.Delete(path);
    }
    internal (List<string> Deleted,List<string> Failed) DeleteHistoryBatch(IEnumerable<string> ids)
    {
        // Validate the entire batch before deleting anything; never accept arbitrary paths.
        var paths=ids.Distinct(StringComparer.OrdinalIgnoreCase).Select(id=>(Id:id,Path:HistoryPath(id))).ToList();
        Directory.CreateDirectory(Root);using var held=LockHistory();
        var deleted=new List<string>();var failed=new List<string>();
        foreach(var item in paths)
        {
            try{File.Delete(item.Path);deleted.Add(item.Id);}
            catch(Exception ex) when(ex is IOException or UnauthorizedAccessException){failed.Add(item.Id);}
        }
        return(deleted,failed);
    }
    private FileStream LockHistory()
    {
        for(int attempt=0;;attempt++)try{return new(Path.Combine(Root,"history.lock"),FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);}
        catch(IOException) when(attempt<20){Thread.Sleep(25);}
    }
    internal List<EnvelopeHistoryEntry> History(out int unreadable)
    {
        unreadable=0;string folder=Path.Combine(Root,"History");if(!Directory.Exists(folder))return [];
        var rows=new List<EnvelopeHistoryEntry>();
        foreach(string path in Directory.EnumerateFiles(folder,"*.json"))
            try{var entry=Read<EnvelopeHistoryEntry>(path);if(entry is not null){ValidateEntry(entry,Path.GetFileNameWithoutExtension(path));rows.Add(entry);}else unreadable++;}
            catch(Exception ex) when(ex is IOException or InvalidDataException or JsonException){unreadable++;}
        return rows.OrderByDescending(e=>e.RequestedUtc).ToList();
    }
    private string HistoryPath(string id)
    {if(!Guid.TryParseExact(id,"N",out _))throw new InvalidDataException("Invalid history identifier.");return Path.Combine(Root,"History",id+".json");}
    private static T? Read<T>(string path)=>Settings.SettingsDatabase.Read(path) is {} json?JsonSerializer.Deserialize<T>(json):default;
    private static void Write<T>(string path,T value)
    {
        Settings.SettingsDatabase.Write(path,JsonSerializer.Serialize(value,Json));
    }
}
