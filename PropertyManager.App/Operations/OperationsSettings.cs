using System.IO;
using System.Net.Mail;
using System.Text.Json;
using PropertyManager.App.BuildingLink;
using PropertyManager.App.Services;

namespace PropertyManager.App.Operations;

internal sealed class OperationsSettings
{
    public string Workbook {get;set;}="";
    internal const string SharedProperty="Example Condominium - DEMO 1001";
    public string PropertyName {get=>SharedProperty;set{ /* Normalize older settings to the shared calendar. */ }}
    public int DaysAhead {get;set;}=7;
    public bool AllowPartial {get;set;}
    public string SmtpServer {get;set;}="smtp.example.invalid";
    public int SmtpPort {get;set;}=587;
    public string From {get;set;}="";
    public List<string> Recipients {get;set;}=[];
    public List<string> Cc {get;set;}=[];
    public string Subject {get;set;}="Property Operations Daily Reminder";
    public string Time {get;set;}="08:00";
    public List<int> Days {get;set;}=[1,2,3,4,5];
    internal static string DirectoryPath=>Path.Combine(AppSettings.DataDirectory,"OperationsReminder");
    internal static string SettingsPath=>Path.Combine(DirectoryPath,"settings.json");
    internal static OperationsSettings Load()=>Settings.SettingsDatabase.Read(SettingsPath) is {} json?JsonSerializer.Deserialize<OperationsSettings>(json)??throw new InvalidDataException("Reminder settings are empty."):new();
    internal void Save(){Validate();Directory.CreateDirectory(DirectoryPath);BuildingLinkSettings.AtomicWrite(SettingsPath,JsonSerializer.Serialize(this));}
    internal void Validate(bool sending=false)
    {
        if(!Path.IsPathFullyQualified(Workbook)||!Path.GetExtension(Workbook).Equals(".xlsx",StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("Choose the full path to your .xlsx operations workbook.");
        if(DaysAhead<0||DaysAhead>90)throw new InvalidDataException("Set days ahead between 0 and 90.");
        if(sending)ValidateEmail();
    }
    internal void SaveEmailSettings(bool accountOnly=false)
    {
        if(accountOnly){ValidateSender();foreach(string address in Recipients.Concat(Cc))if(!MailAddress.TryCreate(address,out var parsed)||parsed.Address!=address)throw new InvalidDataException("Use complete reminder recipient addresses.");if(Subject.Contains('\r')||Subject.Contains('\n'))throw new InvalidDataException("The subject must be one line.");}else ValidateEmail();Directory.CreateDirectory(DirectoryPath);BuildingLinkSettings.AtomicWrite(SettingsPath,JsonSerializer.Serialize(this));
    }
    internal void ValidateSender()
    {
        if(string.IsNullOrWhiteSpace(SmtpServer)||SmtpServer.Any(char.IsWhiteSpace)||SmtpPort<1||SmtpPort>65535||!MailAddress.TryCreate(From,out var parsed)||parsed.Address!=From)throw new InvalidDataException("Set a valid email server, port and From address in Manage credentials > Email sending.");
    }
    internal void ValidateEmail()
    {
        if(string.IsNullOrWhiteSpace(SmtpServer)||SmtpServer.Any(char.IsWhiteSpace)||SmtpPort<1||SmtpPort>65535||Recipients.Count==0||string.IsNullOrWhiteSpace(Subject))throw new InvalidDataException("Set the email server, port (1-65535), recipients and subject in Manage credentials > Email sending.");
        foreach(string value in Recipients.Concat(Cc).Append(From))
            if(!MailAddress.TryCreate(value,out var parsed)||parsed.Address!=value)throw new InvalidDataException("Use complete email addresses without display names.");
        if(Subject.Contains('\r')||Subject.Contains('\n'))throw new InvalidDataException("The subject must be one line.");
    }
    internal static List<string> Split(string text)=>text.Split([',',';','\n','\r'],StringSplitOptions.RemoveEmptyEntries|StringSplitOptions.TrimEntries).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    internal static OperationsSettings Import(string path)
    {
        using var json=JsonDocument.Parse(File.ReadAllText(path));var r=json.RootElement;var value=new OperationsSettings();
        string Read(JsonElement e,string key,string fallback="")=>e.TryGetProperty(key,out var item)&&item.ValueKind==JsonValueKind.String?item.GetString()??fallback:fallback;
        List<string> Strings(JsonElement e,string key,List<string> fallback)=>e.TryGetProperty(key,out var item)&&item.ValueKind==JsonValueKind.Array?item.EnumerateArray().Select(v=>v.GetString()??"").Where(v=>v.Length>0).ToList():fallback;
        value.Workbook=Read(r,"workbook");value.Subject=Read(r,"subject",value.Subject);
        if(r.TryGetProperty("buildinglink",out var b)){value.PropertyName=Read(b,"property_name",value.PropertyName);if(b.TryGetProperty("days_ahead",out var d))value.DaysAhead=d.GetInt32();}
        if(r.TryGetProperty("email",out var e)){value.SmtpServer=Read(e,"smtp_server",value.SmtpServer);value.From=Read(e,"from");value.Recipients=Strings(e,"to",[]);value.Cc=Strings(e,"cc",[]);if(e.TryGetProperty("smtp_port",out var port))value.SmtpPort=port.GetInt32();}
        // Credentials are deliberately never imported from config files.
        return value;
    }
}
