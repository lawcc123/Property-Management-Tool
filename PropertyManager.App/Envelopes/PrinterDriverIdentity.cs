using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;

namespace PropertyManager.App.Envelopes;

internal static class PrinterDriverIdentity
{
    [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)] private struct DriverInfo
    {public uint Version;public IntPtr Name,Environment,DriverPath,DataFile,ConfigFile;}
    internal static string Read(string printer)
    {
        if(!OpenPrinter(printer,out var handle,IntPtr.Zero))throw new Win32Exception(Marshal.GetLastWin32Error());
        IntPtr buffer=IntPtr.Zero;
        try
        {
            GetPrinterDriver(handle,null,2,IntPtr.Zero,0,out uint size);if(size==0)throw new Win32Exception(Marshal.GetLastWin32Error());
            buffer=Marshal.AllocHGlobal(checked((int)size));if(!GetPrinterDriver(handle,null,2,buffer,size,out _))throw new Win32Exception(Marshal.GetLastWin32Error());
            var driver=Marshal.PtrToStructure<DriverInfo>(buffer);
            string FileIdentity(IntPtr pointer)
            {
                string path=Marshal.PtrToStringUni(pointer)??"";if(!File.Exists(path))throw new InvalidDataException("The installed printer driver files could not be verified. Reinstall or update the driver before saving a profile.");
                var file=new FileInfo(path);return path+"|"+FileVersionInfo.GetVersionInfo(path).FileVersion+"|"+file.Length+"|"+file.LastWriteTimeUtc.Ticks;
            }
            return string.Join(";",driver.Version,Marshal.PtrToStringUni(driver.Name),Marshal.PtrToStringUni(driver.Environment),FileIdentity(driver.DriverPath),FileIdentity(driver.DataFile),FileIdentity(driver.ConfigFile));
        }
        finally{if(buffer!=IntPtr.Zero)Marshal.FreeHGlobal(buffer);ClosePrinter(handle);}
    }
    [DllImport("winspool.drv",EntryPoint="OpenPrinterW",CharSet=CharSet.Unicode,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] private static extern bool OpenPrinter(string name,out IntPtr printer,IntPtr defaults);
    [DllImport("winspool.drv",SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] private static extern bool ClosePrinter(IntPtr printer);
    [DllImport("winspool.drv",EntryPoint="GetPrinterDriverW",CharSet=CharSet.Unicode,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] private static extern bool GetPrinterDriver(IntPtr printer,string? environment,uint level,IntPtr info,uint bytes,out uint needed);
}
