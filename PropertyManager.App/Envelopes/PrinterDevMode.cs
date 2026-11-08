using System.Buffers.Binary;
using System.IO;

namespace PropertyManager.App.Envelopes;

// Unicode DEVMODE public fields only. Never interpret or patch private driver bytes.
internal static class PrinterDevMode
{
    private const uint Orientation=0x1,PaperSize=0x2,PaperLength=0x4,PaperWidth=0x8,Copies=0x100,DefaultSource=0x200,Duplex=0x1000,FormName=0x10000;
    private const uint Layout=Orientation|PaperSize|PaperLength|PaperWidth|FormName;
    private static ushort U16(byte[] mode,int offset)=>BinaryPrimitives.ReadUInt16LittleEndian(mode.AsSpan(offset,2));
    private static uint U32(byte[] mode,int offset)=>BinaryPrimitives.ReadUInt32LittleEndian(mode.AsSpan(offset,4));
    private static void Put16(byte[] mode,int offset,ushort value)=>BinaryPrimitives.WriteUInt16LittleEndian(mode.AsSpan(offset,2),value);
    private static void Check(byte[] mode)
    {
        if(mode.Length<96)throw new InvalidDataException("The printer returned an incomplete native settings buffer.");
        int size=U16(mode,68),extra=U16(mode,70);
        if(size<96||size+extra>mode.Length)throw new InvalidDataException("The printer returned invalid native settings lengths.");
    }
    internal static byte[] PreserveSelection(byte[] mode,byte[] requested,int trayId)
    {
        Check(mode);Check(requested);
        if(trayId<0||trayId>ushort.MaxValue)throw new InvalidDataException("The printer tray ID is invalid.");
        byte[] result=mode.ToArray();uint fields=(U32(result,72)&~Layout)|(U32(requested,72)&Layout)|DefaultSource|Duplex|Copies;
        foreach(var (flag,offset) in new[]{(Orientation,76),(PaperSize,78),(PaperLength,80),(PaperWidth,82)})
            if((fields&flag)!=0)requested.AsSpan(offset,2).CopyTo(result.AsSpan(offset,2));
        if((fields&FormName)!=0)
        {
            if(U16(requested,68)<166||U16(result,68)<166)throw new InvalidDataException("The printer returned an incomplete paper-form name.");
            requested.AsSpan(102,64).CopyTo(result.AsSpan(102,64));
        }
        Put16(result,88,(ushort)trayId);Put16(result,94,1);Put16(result,86,1);
        BinaryPrimitives.WriteUInt32LittleEndian(result.AsSpan(72,4),fields);return result;
    }
    internal static string PaperDescription(byte[] mode)
    {
        Check(mode);uint fields=U32(mode,72);
        string id=(fields&PaperSize)!=0?U16(mode,78).ToString():"not reported";
        string form=(fields&FormName)!=0&&U16(mode,68)>=166
            ?System.Text.Encoding.Unicode.GetString(mode,102,64).TrimEnd('\0'):"not reported";
        string dimensions=(fields&(PaperWidth|PaperLength))==(PaperWidth|PaperLength)
            ?$"; native dimensions {U16(mode,82)/10.0:0.###} x {U16(mode,80)/10.0:0.###} mm":"";
        return $"ID {id}{(id=="20"?" (COM10 / #10)":"")}; form {form}{dimensions}";
    }
    internal static byte[] CustomEnvelopePage(byte[] mode,double widthMm,double heightMm,bool pageAxes=false)
    {
        Check(mode);
        if(!double.IsFinite(widthMm)||!double.IsFinite(heightMm)||widthMm<50||heightMm<50||widthMm>450||heightMm>450)
            throw new InvalidDataException("Custom envelope dimensions are invalid.");
        var result=mode.ToArray();
        // DEVMODE public fields: dmPaperSize MUST be zero when explicit
        // dimensions are used. Preserve all private driver/output/media bytes.
        uint fields=(U32(result,72)&~FormName)|Orientation|PaperSize|PaperLength|PaperWidth|0x10;
        Put16(result,76,(ushort)(!pageAxes&&widthMm>heightMm?2:1));
        Put16(result,78,0);
        Put16(result,80,checked((ushort)Math.Round((pageAxes?heightMm:Math.Max(widthMm,heightMm))*10)));
        Put16(result,82,checked((ushort)Math.Round((pageAxes?widthMm:Math.Min(widthMm,heightMm))*10)));
        Put16(result,84,100); // Actual-size document; never fit to Letter.
        if(U16(result,68)>=166)result.AsSpan(102,64).Clear();
        BinaryPrimitives.WriteUInt32LittleEndian(result.AsSpan(72,4),fields);
        return result;
    }
    internal static int? Tray(byte[] mode){Check(mode);return(U32(mode,72)&DefaultSource)!=0?U16(mode,88):null;}
    internal static int CopiesValue(byte[] mode){Check(mode);return(U32(mode,72)&Copies)!=0?U16(mode,86):1;}
    internal static int? DuplexValue(byte[] mode){Check(mode);return(U32(mode,72)&Duplex)!=0?U16(mode,94):null;}
    internal static void RequireProfile(byte[] mode,EnvelopeOptions options)
    {
        Check(mode);
        if(Tray(mode)!=options.TrayId||options.TrayId is null)throw new InvalidDataException("The driver profile does not select '"+options.TrayName+"'. Choose that tray in Set up print profile.");
        if(DuplexValue(mode)!=1||U16(mode,86)!=1)throw new InvalidDataException("Set the driver profile to single-sided printing and one copy. The app controls mailing copy counts.");
    }
    internal static void RequireTransfer(byte[] expected,byte[] actual,EnvelopeOptions options)
    {
        RequireProfile(expected,options);RequireProfile(actual,options);
        uint mask=Layout|DefaultSource|Duplex|Copies|0x2000000; // includes native media type
        if((U32(expected,72)&mask)!=(U32(actual,72)&mask))throw new InvalidDataException("The driver changed the saved paper/tray settings while preparing submission. Set up the print profile again. No print job was sent.");
        foreach(var(flag,offset) in new[]{(Orientation,76),(PaperSize,78),(PaperLength,80),(PaperWidth,82)})
            if((U32(expected,72)&flag)!=0&&U16(expected,offset)!=U16(actual,offset))throw new InvalidDataException("The driver changed the saved paper size or orientation. No print job was sent.");
        if((U32(expected,72)&FormName)!=0&&!expected.AsSpan(102,64).SequenceEqual(actual.AsSpan(102,64)))throw new InvalidDataException("The driver changed the saved paper form. No print job was sent.");
        if((U32(expected,72)&0x2000000)!=0&&(U16(expected,68)<200||U16(actual,68)<200||U32(expected,196)!=U32(actual,196)))throw new InvalidDataException("The driver changed the saved paper type. No print job was sent.");
        int expectedSize=U16(expected,68),actualSize=U16(actual,68),extra=U16(expected,70);
        if(extra!=U16(actual,70)||!expected.AsSpan(expectedSize,extra).SequenceEqual(actual.AsSpan(actualSize,extra)))throw new InvalidDataException("The driver changed its saved private tray/media settings during submission setup. Set up the print profile again. No print job was sent.");
    }
    internal static byte[] FinalizeSelection(byte[] converted,byte[] requested,TrayChoice source,Func<byte[],byte[]> driverValidate,Func<int,string> trayName,bool canDuplex,string paper)
    {
        // PrintTicket conversion can recreate Auto Select and drop a driver's
        // custom input-bin ID. Reapply explicit public requests, then let the
        // driver reconcile its private data before checking the accepted result.
        byte[] native=driverValidate(PreserveSelection(converted,requested,source.Id));
        int? tray=Tray(native),duplex=DuplexValue(native);
        if(tray!=source.Id)
            throw new InvalidOperationException("The printer did not accept '"+source.Name+"' (ID "+source.Id+") for "+paper+". It returned "+(tray.HasValue?"'"+trayName(tray.Value)+"' (ID "+tray.Value+")":"no confirmed tray")+". No print job was sent. Check that this tray supports the selected paper size and type in Driver preferences.");
        if(duplex!=1&&(duplex.HasValue||canDuplex))
            throw new InvalidOperationException("The printer did not accept single-sided printing for "+paper+". It returned "+(duplex==2?"two-sided (long edge)":duplex==3?"two-sided (short edge)":"an unconfirmed duplex setting")+". No print job was sent. Set single-sided printing in Driver preferences.");
        return native;
    }
}
