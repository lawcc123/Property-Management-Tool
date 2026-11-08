using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace PropertyManager.App.Envelopes;

internal sealed record NativePrinterProfile(string Printer,string Driver,string Machine,int Architecture,string Media,double WidthMm,double HeightMm,int TrayId,string TrayName,byte[] Mode);
internal sealed class PrinterDriverProfiles(string? directory=null)
{
    private readonly string root=directory??Path.Combine(Services.AppSettings.DataDirectory,"EnvelopePrinter","DriverProfiles");
    private static readonly byte[] Entropy=Encoding.UTF8.GetBytes("PropertyManagerDemo.NativePrinterProfile.v1");
    internal static string Key(EnvelopeOptions o)=>Convert.ToHexString(SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(new{o.Printer,Media=o.IsLabel?"Labels":"Envelope",o.WidthMm,o.HeightMm})));
    internal void Save(EnvelopeOptions options,string driver,byte[] mode)
    {
        options.Validate();PrinterDevMode.RequireProfile(mode,options);
        var record=new NativePrinterProfile(options.Printer,driver,Environment.MachineName,IntPtr.Size,options.IsLabel?"Labels":"Envelope",options.WidthMm,options.HeightMm,options.TrayId!.Value,options.TrayName,mode.ToArray());
        byte[] plain=JsonSerializer.SerializeToUtf8Bytes(record);
        try
        {
            byte[] secured=ProtectedData.Protect(plain,Entropy,DataProtectionScope.CurrentUser);Directory.CreateDirectory(root);
            string path=Path.Combine(root,Key(options)+".dpapi"),temp=path+"."+Guid.NewGuid().ToString("N")+".tmp";
            try{File.WriteAllBytes(temp,secured);File.Move(temp,path,true);}finally{if(File.Exists(temp))File.Delete(temp);}
        }
        finally{CryptographicOperations.ZeroMemory(plain);CryptographicOperations.ZeroMemory(record.Mode);}
    }
    internal byte[]? Load(EnvelopeOptions options,string driver)
    {
        string path=Path.Combine(root,Key(options)+".dpapi");if(!File.Exists(path))return null;
        byte[] plain=ProtectedData.Unprotect(File.ReadAllBytes(path),Entropy,DataProtectionScope.CurrentUser);
        try
        {
            var record=JsonSerializer.Deserialize<NativePrinterProfile>(plain)??throw new InvalidDataException("The saved driver profile is empty.");
            try
            {
                if(record.Printer!=options.Printer||record.Driver!=driver||record.Machine!=Environment.MachineName||record.Architecture!=IntPtr.Size||record.Media!=(options.IsLabel?"Labels":"Envelope")||record.WidthMm!=options.WidthMm||record.HeightMm!=options.HeightMm||record.TrayId!=options.TrayId||record.TrayName!=options.TrayName)
                    throw new InvalidDataException("The saved driver profile does not match this printer, driver, tray or mailing size. Set up the print profile again.");
                PrinterDevMode.RequireProfile(record.Mode,options);return record.Mode.ToArray();
            }
            finally{CryptographicOperations.ZeroMemory(record.Mode);}
        }
        finally{CryptographicOperations.ZeroMemory(plain);}
    }
    internal void Delete(EnvelopeOptions options)=>File.Delete(Path.Combine(root,Key(options)+".dpapi"));
}
