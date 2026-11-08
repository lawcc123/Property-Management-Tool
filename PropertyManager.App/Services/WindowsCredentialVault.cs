using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PropertyManager.App.Services;

internal static class WindowsCredentialVault
{
    private const uint Generic = 1;
    private const uint LocalMachine = 2;
    private const int NotFound = 1168;

    private static string Target(string target) => target;

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
        throw new Win32Exception(error, "Windows could not read the saved service sign-in.");
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
                Comment = "Saved app login for this Windows user."
            };
            if (!CredWrite(ref item, 0)) throw new Win32Exception(Marshal.GetLastWin32Error(), "Windows could not save the service sign-in.");
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
        if (error != NotFound) throw new Win32Exception(error, "Windows could not remove the saved service sign-in.");
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

internal static class BuildingLinkCredentialStore
{
    internal const string Target = "PropertyManagerDemo/BuildingLink";
}
