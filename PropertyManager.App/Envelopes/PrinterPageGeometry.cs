using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;

namespace PropertyManager.App.Envelopes;

// GDI caps already use the final DEVMODE orientation. Do not rotate them again.
internal sealed record PrinterPageGeometry(Rect Page,Rect Imageable)
{
    internal string Description=>"Driver page: "+Mm(Page)+"; printable area: "+Mm(Imageable)+" (left, top, width, height in mm).";
    internal static string Mm(Rect rect)=>string.Join(", ",new[]{rect.X,rect.Y,rect.Width,rect.Height}.Select(v=>(v/EnvelopeRenderer.DipsPerMm).ToString("0.##",System.Globalization.CultureInfo.InvariantCulture)));
    internal void RequireSize(double widthMm,double heightMm)
    {
        if(Math.Abs(Page.Width/EnvelopeRenderer.DipsPerMm-widthMm)>0.8||Math.Abs(Page.Height/EnvelopeRenderer.DipsPerMm-heightMm)>0.8)
            throw new InvalidOperationException("The printer's actual page size/orientation does not match the preview. "+Description+" No print job was sent. Check paper size and orientation in Driver preferences.");
    }
    internal static PrinterPageGeometry FromCaps(Func<int,int> get)
    {
        int dpiX=get(88),dpiY=get(90),physicalWidth=get(110),physicalHeight=get(111),offsetX=get(112),offsetY=get(113),width=get(8),height=get(10);
        if(dpiX<=0||dpiY<=0||physicalWidth<=0||physicalHeight<=0||width<=0||height<=0||offsetX<0||offsetY<0||(long)offsetX+width>physicalWidth+1L||(long)offsetY+height>physicalHeight+1L)
            throw new InvalidDataException("The printer returned invalid page or printable-area measurements. No print job was sent. Check its paper size and driver settings.");
        return new(new Rect(0,0,physicalWidth*96.0/dpiX,physicalHeight*96.0/dpiY),new Rect(offsetX*96.0/dpiX,offsetY*96.0/dpiY,width*96.0/dpiX,height*96.0/dpiY));
    }
    internal static PrinterPageGeometry Read(string printer,byte[] devmode)
    {
        // An information context reads driver metrics only; no StartDoc/print job.
        if(devmode.Length<96)throw new InvalidDataException("The printer returned incomplete native page settings.");
        IntPtr input=Marshal.AllocHGlobal(devmode.Length),context=IntPtr.Zero;
        try
        {
            Marshal.Copy(devmode,0,input,devmode.Length);context=CreateIC("WINSPOOL",printer,null,input);
            if(context==IntPtr.Zero)throw new Win32Exception(Marshal.GetLastWin32Error(),"Could not measure this printer's page. No print job was sent.");
            return FromCaps(index=>GetDeviceCaps(context,index));
        }
        finally{if(context!=IntPtr.Zero)DeleteDC(context);Marshal.FreeHGlobal(input);}
    }
    [DllImport("gdi32.dll",EntryPoint="CreateICW",CharSet=CharSet.Unicode,SetLastError=true)] private static extern IntPtr CreateIC(string driver,string device,string? output,IntPtr devmode);
    [DllImport("gdi32.dll")] private static extern int GetDeviceCaps(IntPtr context,int index);
    [DllImport("gdi32.dll")] [return:MarshalAs(UnmanagedType.Bool)] private static extern bool DeleteDC(IntPtr context);
}
