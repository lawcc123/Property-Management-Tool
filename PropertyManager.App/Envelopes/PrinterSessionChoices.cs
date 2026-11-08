using System.Text.Json;
using System.Security.Cryptography;

namespace PropertyManager.App.Envelopes;

internal sealed class PrinterSessionChoices
{
    private string key="";
    private byte[]? mode;
    private static string Key(EnvelopeOptions o)=>JsonSerializer.Serialize(new{o.Printer,o.WidthMm,o.HeightMm,o.IsLabel,o.DefaultToBypass,o.UsesWordEnvelopeCompatibility});
    internal byte[]? Read(EnvelopeOptions options)=>key==Key(options)?mode?.ToArray():null;
    internal void Accept(EnvelopeOptions options,byte[] selected){Clear();key=Key(options);mode=selected.ToArray();}
    internal void Clear(){if(mode is not null)CryptographicOperations.ZeroMemory(mode);mode=null;key="";}
    internal static TrayChoice? InitialTray(IEnumerable<TrayChoice> trays,EnvelopeOptions options,int? driverDefault)
    {
        var all=trays.ToList();
        if(options.TrayId.HasValue)return all.FirstOrDefault(t=>t.Id==options.TrayId&&t.Name.Equals(options.TrayName,StringComparison.OrdinalIgnoreCase));
        return(options.DefaultToBypass?all.FirstOrDefault(t=>t.Name.Contains("Bypass",StringComparison.OrdinalIgnoreCase)||t.Name.Contains("By-pass",StringComparison.OrdinalIgnoreCase)):null)??all.FirstOrDefault(t=>t.Id==driverDefault);
    }
}
internal interface IEnvelopePrinterProperties
{
    bool DriverPreferences(EnvelopeOptions options,IntPtr owner);
    TrayChoice? SelectedTray {get;}
    void ClearDriverPreferences();
}
