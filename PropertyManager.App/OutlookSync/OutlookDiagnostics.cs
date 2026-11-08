using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace PropertyManager.App.OutlookSync;

internal static class OutlookDiagnostics
{
    internal static string PathName=>Path.Combine(Services.AppSettings.DataDirectory,"OutlookSync","last-error.dpapi");
    private static readonly byte[] Entropy=Encoding.UTF8.GetBytes("PropertyManagerDemo.Outlook.Diagnostics.v1");
    internal static string Record(string operation,string context,Exception error)
    {
        List<object> chain=[];
        for(Exception? current=error;current is not null;current=current.InnerException)
            chain.Add(new{type=current.GetType().FullName,message=current.Message,hresult="0x"+current.HResult.ToString("X8")});
        string json=JsonSerializer.Serialize(new{time=DateTimeOffset.UtcNow,operation,context,appVersion=Assembly.GetExecutingAssembly().GetName().Version?.ToString(),windows=Environment.OSVersion.VersionString,process64Bit=Environment.Is64BitProcess,dotnet=RuntimeInformation.FrameworkDescription,errors=chain,stack=error.ToString()},new JsonSerializerOptions{WriteIndented=true});
        try
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(PathName)!);
            byte[] bytes=ProtectedData.Protect(Encoding.UTF8.GetBytes(json),Entropy,DataProtectionScope.CurrentUser);
            string temp=PathName+"."+Guid.NewGuid().ToString("N")+".tmp";
            try{File.WriteAllBytes(temp,bytes);File.Move(temp,PathName,true);}finally{if(File.Exists(temp))File.Delete(temp);}
        }
        catch(Exception ex)when(ex is IOException or UnauthorizedAccessException or CryptographicException){/* Keep an in-memory report even if disk/vault is unavailable. */}
        return json;
    }
    // Export only diagnostic fields; never export the raw stack or an arbitrary old payload.
    internal static string ExportReport(string stored)
    {
        using var document=JsonDocument.Parse(stored);
        var root=document.RootElement;
        string Field(JsonElement source,string name)=>source.TryGetProperty(name,out var value)?Redact(value.ToString()):"";
        var errors=new List<object>();
        if(root.TryGetProperty("errors",out var chain)&&chain.ValueKind==JsonValueKind.Array)
            foreach(var item in chain.EnumerateArray())
                if(item.ValueKind==JsonValueKind.Object)errors.Add(new{type=Field(item,"type"),hresult=Field(item,"hresult"),message=Field(item,"message")});
        return JsonSerializer.Serialize(new{time=Field(root,"time"),operation=Field(root,"operation"),context=Field(root,"context"),appVersion=Field(root,"appVersion"),windows=Field(root,"windows"),process64Bit=Field(root,"process64Bit"),dotnet=Field(root,"dotnet"),errors,privacy="Raw stack omitted; recognizable email addresses, Windows paths and credential values redacted. Review before sharing."},new JsonSerializerOptions{WriteIndented=true});
    }
    private static string Redact(string value)
    {
        value=System.Text.RegularExpressions.Regex.Replace(value,@"(?i)\bBearer\s+[^\s,;]+","Bearer [redacted]");
        value=System.Text.RegularExpressions.Regex.Replace(value,@"(?i)\b(authorization|cookie)\b\s*[:=][^\r\n]*","$1=[redacted]");
        value=System.Text.RegularExpressions.Regex.Replace(value,@"(?i)\b(password|passwd|pwd|access_token|refresh_token|client_secret)\b[""']?\s*[:=]\s*(?:""[^""]*""|'[^']*'|[^\s,;]+)","$1=[redacted]");
        value=System.Text.RegularExpressions.Regex.Replace(value,@"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}","[email redacted]");
        return System.Text.RegularExpressions.Regex.Replace(value,@"(?i)(?:[a-z]:\\|\\\\)[^\r\n|<>]*","[path redacted]");
    }
    internal static string? Load()=>File.Exists(PathName)?Encoding.UTF8.GetString(ProtectedData.Unprotect(File.ReadAllBytes(PathName),Entropy,DataProtectionScope.CurrentUser)):null;
}
