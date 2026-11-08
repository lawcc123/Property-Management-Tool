using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PropertyManager.App.Services;

public sealed class LoginCredentials
{
    [JsonPropertyName("origin")] public string Origin { get; init; } = "";
    [JsonPropertyName("username")] public string Username { get; init; } = "";
    [JsonPropertyName("password")] public string Password { get; init; } = "";
    public override string ToString() => "Yardi sign-in details (redacted)";
}

/// <summary>Per-user Windows Credential Manager; never writes credentials into app files.</summary>
public static class YardiCredentialStore
{
    private const uint Generic = 1;
    private const uint LocalMachine = 2;
    private const int NotFound = 1168;

    public static string ConfiguredOrigin()
    {
        using var document = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Backend", "p2p_private_config.example.json")));
        var uri = new Uri(document.RootElement.GetProperty("login_url").GetString()!);
        if (uri.Scheme != Uri.UriSchemeHttps) throw new InvalidOperationException("Saved sign-in requires an HTTPS Yardi login address.");
        return uri.GetLeftPart(UriPartial.Authority).ToLowerInvariant();
    }

    private static string Target(string origin) => "PropertyManagerDemo/Yardi/" + new Uri(origin).GetLeftPart(UriPartial.Authority).ToLowerInvariant();

    public static string? Username(string origin)
    {
        if (!ReadPointer(origin, out var pointer)) return null;
        try { return Marshal.PtrToStructure<Credential>(pointer).UserName; }
        finally { CredFree(pointer); }
    }

    public static LoginCredentials? Read(string origin)
    {
        if (!ReadPointer(origin, out var pointer)) return null;
        byte[]? bytes = null;
        try
        {
            var item = Marshal.PtrToStructure<Credential>(pointer);
            bytes = new byte[checked((int)item.CredentialBlobSize)];
            Marshal.Copy(item.CredentialBlob, bytes, 0, bytes.Length);
            return new LoginCredentials { Origin = origin, Username = item.UserName, Password = Encoding.Unicode.GetString(bytes) };
        }
        finally
        {
            if (bytes is not null) CryptographicOperations.ZeroMemory(bytes);
            CredFree(pointer);
        }
    }

    private static bool ReadPointer(string origin, out IntPtr pointer)
    {
        if (CredRead(Target(origin), Generic, 0, out pointer)) return true;
        var error = Marshal.GetLastWin32Error();
        if (error == NotFound) return false;
        throw new Win32Exception(error, "Windows could not read the saved Yardi sign-in.");
    }

    public static void Save(LoginCredentials login)
    {
        var bytes = Encoding.Unicode.GetBytes(login.Password);
        if (bytes.Length > 2560)
        {
            CryptographicOperations.ZeroMemory(bytes);
            throw new InvalidOperationException("The password is too long for Windows Credential Manager. Use manual sign-in.");
        }
        var blob = Marshal.AllocCoTaskMem(bytes.Length);
        try
        {
            Marshal.Copy(bytes, 0, blob, bytes.Length);
            var item = new Credential
            {
                Type = Generic, TargetName = Target(login.Origin), UserName = login.Username,
                CredentialBlob = blob, CredentialBlobSize = (uint)bytes.Length, Persist = LocalMachine,
                Comment = "Optional Yardi autofill for this Windows user."
            };
            if (!CredWrite(ref item, 0)) throw new Win32Exception(Marshal.GetLastWin32Error(), "Windows could not save the Yardi sign-in.");
        }
        finally
        {
            for (var i = 0; i < bytes.Length; i++) Marshal.WriteByte(blob, i, 0);
            Marshal.FreeCoTaskMem(blob);
            CryptographicOperations.ZeroMemory(bytes);
        }
    }

    public static void Delete(string origin)
    {
        if (CredDelete(Target(origin), Generic, 0)) return;
        var error = Marshal.GetLastWin32Error();
        if (error != NotFound) throw new Win32Exception(error, "Windows could not remove the saved Yardi sign-in.");
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct Credential
    {
        public uint Flags;
        public uint Type;
        [MarshalAs(UnmanagedType.LPWStr)] public string TargetName;
        [MarshalAs(UnmanagedType.LPWStr)] public string? Comment;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
        public uint CredentialBlobSize;
        public IntPtr CredentialBlob;
        public uint Persist;
        public uint AttributeCount;
        public IntPtr Attributes;
        [MarshalAs(UnmanagedType.LPWStr)] public string? TargetAlias;
        [MarshalAs(UnmanagedType.LPWStr)] public string UserName;
    }

    [DllImport("advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool CredRead(string target, uint type, uint flags, out IntPtr credential);
    [DllImport("advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool CredWrite(ref Credential credential, uint flags);
    [DllImport("advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)] private static extern bool CredDelete(string target, uint type, uint flags);
    [DllImport("advapi32.dll")] private static extern void CredFree(IntPtr buffer);
}
