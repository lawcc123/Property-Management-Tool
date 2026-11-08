using System.IO;
using System.Security.Cryptography;
using System.Text;
namespace PropertyManager.App.PersonalizedEmail;
internal static class CampaignSignatureSettings
{
    internal const string Default="Management Office\nPropertyManagerDemo\n\nPlease do not reply to this email. For assistance, please contact the Management Office.";
    internal static string PathName=>Path.Combine(Services.AppSettings.DataDirectory,"PersonalizedEmail","signature.dpapi");
    private static readonly byte[] Entropy=Encoding.UTF8.GetBytes("PropertyManagerDemo.PersonalizedEmail.Signature.v1");
    internal static string Load()=>File.Exists(PathName)?Encoding.UTF8.GetString(ProtectedData.Unprotect(File.ReadAllBytes(PathName),Entropy,DataProtectionScope.CurrentUser)):Default;
    internal static void Save(string signature)
    {
        if(signature.Length>20000)throw new InvalidDataException("Keep the signature under 20,000 characters.");
        Directory.CreateDirectory(Path.GetDirectoryName(PathName)!);string stage=PathName+"."+Guid.NewGuid().ToString("N")+".tmp";
        try{File.WriteAllBytes(stage,ProtectedData.Protect(Encoding.UTF8.GetBytes(signature),Entropy,DataProtectionScope.CurrentUser));File.Move(stage,PathName,true);}
        finally{if(File.Exists(stage))File.Delete(stage);}
    }
}
