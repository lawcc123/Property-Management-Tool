using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
namespace PropertyManager.App.PersonalizedEmail;
internal sealed record EmailSignature(string Id,string Name,string Html)
{
    [System.Text.Json.Serialization.JsonIgnore] public string Text=>SignatureHtml.PlainText(Html);
}
internal sealed class SignatureLibrary
{
    public string DefaultId{get;set;}="no-reply";
    public List<EmailSignature> Items{get;set;}=[];
    internal static EmailSignature BuiltIn=>new("no-reply","No-reply (built-in)",SignatureHtml.DefaultHtml);
    internal EmailSignature Default=>All.FirstOrDefault(s=>s.Id==DefaultId)??BuiltIn;
    internal List<EmailSignature> All=>new[]{BuiltIn}.Concat(Items).ToList();
    internal void Validate()
    {
        if(Items.Count>50||Items.Select(s=>s.Id).Distinct().Count()!=Items.Count||Items.Select(s=>s.Name.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).Count()!=Items.Count)throw new InvalidDataException("Use up to 50 signatures with unique names.");
        foreach(var item in Items){if(!Guid.TryParseExact(item.Id,"N",out _)||string.IsNullOrWhiteSpace(item.Name)||item.Name.Length>80)throw new InvalidDataException("Enter a signature name of 1 to 80 characters.");SignatureHtml.Parse(item.Html);}
        if(All.All(s=>s.Id!=DefaultId))DefaultId="no-reply";
    }
}
internal static class SignatureLibraryStore
{
    internal static string PathName=>Path.Combine(Services.AppSettings.DataDirectory,"PersonalizedEmail","signature-library.dpapi");
    static readonly byte[] Entropy=Encoding.UTF8.GetBytes("PropertyManagerDemo.PersonalizedEmail.SignatureLibrary.v1");
    internal static SignatureLibrary Load(string? path=null)
    {
        if(File.Exists(path??PathName))
        {
            var library=JsonSerializer.Deserialize<SignatureLibrary>(ProtectedData.Unprotect(File.ReadAllBytes(path??PathName),Entropy,DataProtectionScope.CurrentUser))??throw new InvalidDataException("The signature library is empty.");library.Validate();return library;
        }
        var result=new SignatureLibrary();
        if(path is null&&File.Exists(CampaignSignatureSettings.PathName))
        {
            string old=CampaignSignatureSettings.Load();
            if(old!=CampaignSignatureSettings.Default){var migrated=new EmailSignature(Guid.NewGuid().ToString("N"),"Previously saved signature",SignatureHtml.FromPlain(old));result.Items.Add(migrated);result.DefaultId=migrated.Id;}
        }
        return result;
    }
    internal static void Save(SignatureLibrary library,string? path=null)
    {
        library.Validate();byte[] data=JsonSerializer.SerializeToUtf8Bytes(library);if(data.Length>15_000_000)throw new InvalidDataException("The signature library exceeds 15 MB. Reduce the number or size of pictures.");
        path??=PathName;Directory.CreateDirectory(Path.GetDirectoryName(path)!);string stage=path+"."+Guid.NewGuid().ToString("N")+".tmp";
        try{File.WriteAllBytes(stage,ProtectedData.Protect(data,Entropy,DataProtectionScope.CurrentUser));File.Move(stage,path,true);}finally{if(File.Exists(stage))File.Delete(stage);}
    }
}
