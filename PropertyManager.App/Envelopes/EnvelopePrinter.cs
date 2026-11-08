using System.Drawing.Printing;
using System.IO;
using System.Printing;
using System.Printing.Interop;
using System.Runtime.InteropServices;
using System.Windows;

namespace PropertyManager.App.Envelopes;

internal sealed record TrayChoice(int Id,string Name){public override string ToString()=>Name;}
internal sealed record PrinterPaper(int Id,string Name,double WidthMm,double HeightMm);
internal sealed record PrinterInfo(string Name,string Driver,List<TrayChoice> Trays,List<PrinterPaper> Papers,string Notes,int? DefaultTrayId=null);
internal sealed record PrinterProgress(int PagesRendered,int TotalPages,int? JobId,string? Settings=null);
internal sealed record PrinterResult(int PagesRendered,int? JobId,bool Cancelled);
internal sealed class PreparedEnvelopeJob : IDisposable
{
    internal required EnvelopeSnapshot Snapshot {get;init;}
    internal required Rect Imageable {get;init;}
    internal string MediaType {get;init;}="Envelope";
    internal string SettingsDiagnostics {get;init;}="";
    internal string DriverIdentity {get;init;}="";
    internal object? DriverState {get;init;}
    public void Dispose(){if(DriverState is byte[] bytes)System.Security.Cryptography.CryptographicOperations.ZeroMemory(bytes);if(DriverState is IDisposable state)state.Dispose();}
}
internal interface IEnvelopePrinter
{
    Task<List<string>> PrintersAsync();
    Task<PrinterInfo> DescribeAsync(string name);
    Task<PreparedEnvelopeJob> PrepareAsync(EnvelopeSnapshot snapshot);
    Task<PrinterResult> SubmitAsync(PreparedEnvelopeJob job,string documentName,Action<PrinterProgress> progress,CancellationToken cancellation);
    Task<string> QueueStatusAsync(string printer,int jobId,string documentName);
}

internal sealed class PrinterMediaSetupException(string media,Exception inner)
    : InvalidOperationException("The "+media+" printer settings need attention. Open Printer properties and check the tray, paper size and paper type. No print job has been submitted. Details: "+inner.Message,inner)
{
    internal string Media {get;}=media;
}

internal sealed class WindowsEnvelopePrinter : IEnvelopePrinter,IEnvelopePrinterProperties
{
    internal string LastLayoutDiagnostics {get;private set;}="No printer layout check has been performed in this app session.";
    private readonly PrinterSessionChoices choices=new();
    public TrayChoice? SelectedTray {get;private set;}
    public void ClearDriverPreferences(){choices.Clear();SelectedTray=null;}
    public Task<List<string>> PrintersAsync()=>Task.Run(()=>PrinterSettings.InstalledPrinters.Cast<string>().OrderBy(n=>n,StringComparer.OrdinalIgnoreCase).ToList());
    public Task<PrinterInfo> DescribeAsync(string name)=>Task.Run(()=>
    {
        var settings=Settings(name);string driver="Not reported",notes="";
        try{WithQueue(name,q=>{driver=q.QueueDriver.Name;notes=q.IsOffline?"Windows reports this printer offline.":q.IsPaused?"Windows reports this queue paused.":"Tray and paper sizes read from the installed driver.";return 0;});}
        catch(Exception){notes="Queue details are unavailable. Tray and paper sizes are from the installed driver.";}
        return new PrinterInfo(name,driver,settings.PaperSources.Cast<PaperSource>().Select(s=>new TrayChoice(s.RawKind,s.SourceName)).ToList(),
            settings.PaperSizes.Cast<PaperSize>().Select(p=>new PrinterPaper(p.RawKind,p.PaperName,p.Width*0.254,p.Height*0.254)).ToList(),notes,settings.DefaultPageSettings.PaperSource.RawKind);
    });
    private static PrinterSettings Settings(string name)
    {
        if(string.IsNullOrWhiteSpace(name))throw new InvalidOperationException("Select an installed printer.");
        var settings=new PrinterSettings{PrinterName=name};if(!settings.IsValid)throw new InvalidOperationException("This printer is not available. Refresh the printer list.");return settings;
    }
    public bool DriverPreferences(EnvelopeOptions options,IntPtr owner)
    {
        options.Validate();string printer=options.Printer;
        var settings=Settings(printer);var page=settings.DefaultPageSettings;
        byte[]? saved=choices.Read(options);
        if(saved is not null){try{ApplyMode(settings,page,saved);}finally{System.Security.Cryptography.CryptographicOperations.ZeroMemory(saved);}}
        else ConfigureInitialPage(settings,page,options);
        byte[] initial=Mode(settings,page);
        if(!OpenPrinter(printer,out var handle,IntPtr.Zero))throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        IntPtr input=IntPtr.Zero,output=IntPtr.Zero;
        try
        {
            int size=DocumentProperties(owner,handle,printer,IntPtr.Zero,IntPtr.Zero,0);
            if(size<=0)throw new InvalidOperationException("The printer driver did not report a valid settings size.");
            input=Marshal.AllocHGlobal(initial.Length);Marshal.Copy(initial,0,input,initial.Length);
            output=Marshal.AllocHGlobal(size);
            int result=DocumentProperties(owner,handle,printer,output,input,0x2|0x8|0x4);
            if(result<0)throw new InvalidOperationException("The printer driver could not open its preferences.");
            if(result!=1)return false;
            byte[] selected=Bytes(output,size);RequireSingleCopy(selected);choices.Accept(options,selected);
            int? selectedTray=PrinterDevMode.Tray(selected);SelectedTray=settings.PaperSources.Cast<PaperSource>().Where(t=>t.RawKind==selectedTray).Select(t=>new TrayChoice(t.RawKind,t.SourceName)).FirstOrDefault();
            System.Security.Cryptography.CryptographicOperations.ZeroMemory(selected);return true;
        }
        finally
        {
            if(input!=IntPtr.Zero)Marshal.FreeHGlobal(input);
            if(output!=IntPtr.Zero)Marshal.FreeHGlobal(output);
            ClosePrinter(handle);
        }
    }
    private static void ConfigurePage(PrinterSettings settings,PageSettings page,EnvelopeOptions o)
    {
        var selected=PrinterSessionChoices.InitialTray(settings.PaperSources.Cast<PaperSource>().Select(t=>new TrayChoice(t.RawKind,t.SourceName)),o,page.PaperSource.RawKind);
        var source=selected is null?null:settings.PaperSources.Cast<PaperSource>().FirstOrDefault(t=>t.RawKind==selected.Id);
        if(o.UsesWordEnvelopeCompatibility)
        {
            // Driver properties describe the base job (Letter in the working
            // office Word setup). The custom document size is applied separately.
            var original=settings.PaperSizes.Cast<PaperSize>().FirstOrDefault(p=>p.RawKind==1);
            if(original is not null)page.PaperSize=original;
            if(source is not null)page.PaperSource=source;
            settings.Duplex=Duplex.Simplex;settings.Copies=1;settings.PrintToFile=false;return;
        }
        var paper=SelectPaper(settings.PaperSizes.Cast<PaperSize>(),o);
        if(paper is null&&o.Preset!="Custom")throw new InvalidOperationException("The driver does not report this mailing size. Configure the custom size in the printer driver first.");
        paper??=new PaperSize("Mailing custom",(int)Math.Round(o.WidthMm/0.254),(int)Math.Round(o.HeightMm/0.254));
        page.PaperSize=paper;if(source is not null)page.PaperSource=source;page.Landscape=(paper.Width<paper.Height)!=(o.WidthMm<o.HeightMm);page.Margins=new(0,0,0,0);
        settings.Duplex=Duplex.Simplex;settings.Copies=1;settings.PrintToFile=false;
    }
    // A same-size custom form can appear before the driver's named envelope form.
    // Prefer the Windows standard form ID, while still requiring matching dimensions.
    internal static PaperSize? SelectPaper(IEnumerable<PaperSize> papers,EnvelopeOptions o)
    {
        var matches=papers.Where(p=>Fits(p.Width*0.254,p.Height*0.254,o.WidthMm,o.HeightMm)).ToList();
        int standard=o.IsLabel?1:o.Preset switch{"10"=>20,"DL"=>27,"C5"=>28,"C6"=>31,"Monarch"=>37,_=>0};
        return (standard==0?null:matches.FirstOrDefault(p=>p.RawKind==standard))
            ??(o.IsLabel||o.Preset!="10"?null:matches.FirstOrDefault(p=>
                p.PaperName.Contains("#10",StringComparison.OrdinalIgnoreCase)||new[]{"COM10","ENV10","ENVELOPE10","NUMBER10"}.Any(token=>
                    new string(p.PaperName.Where(char.IsLetterOrDigit).ToArray()).Contains(
                        new string(token.Where(char.IsLetterOrDigit).ToArray()),StringComparison.OrdinalIgnoreCase))))
            ??matches.FirstOrDefault();
    }
    private static void ConfigureInitialPage(PrinterSettings settings,PageSettings page,EnvelopeOptions o)
    {
        ConfigurePage(settings,page,o);
        // Reconcile the named form/tray with the driver's private settings through
        // its supported API, rather than carrying private Letter defaults forward.
        byte[] requested=Mode(settings,page);
        byte[]? accepted=null;
        try{accepted=ValidateMode(o.Printer,requested);ApplyMode(settings,page,accepted);}
        finally
        {
            System.Security.Cryptography.CryptographicOperations.ZeroMemory(requested);
            if(accepted is not null)System.Security.Cryptography.CryptographicOperations.ZeroMemory(accepted);
        }
    }
    internal static byte[] DocumentMode(byte[] mode,EnvelopeOptions options)=>options.UsesWordEnvelopeCompatibility
        ?PrinterDevMode.CustomEnvelopePage(mode,options.WidthMm,options.HeightMm):mode;
    // Modified public DEVMODE dimensions must be reconciled by the driver before
    // CreateIC or printing. Raw public patches can retain private Letter geometry.
    internal static (byte[] Mode,PrinterPageGeometry Geometry,string Diagnostics) ResolveEnvelopeMode(
        byte[] initial,EnvelopeOptions options,Func<byte[],byte[]> validate,
        Func<byte[],byte[]> roundTrip,Func<byte[],PrinterPageGeometry> measure)
    {
        var attempts=new List<string>();
        // The successful office Word capture uses wide custom page axes with
        // portrait orientation. Try that before rotating conventional stock axes.
        foreach(string variant in new[]{"Selected driver form","Custom page axes","Custom stock with orientation"})
        {
            byte[] request=variant=="Selected driver form"?initial.ToArray():
                PrinterDevMode.CustomEnvelopePage(initial,options.WidthMm,options.HeightMm,variant=="Custom page axes");
            byte[]? accepted=null,transferred=null,final=null;
            try
            {
                accepted=validate(request);
                RequireSubmissionSelection(initial,accepted);
                var acceptedGeometry=measure(accepted);
                attempts.Add(variant+": "+PrinterDevMode.PaperDescription(accepted)+". "+acceptedGeometry.Description);
                acceptedGeometry.RequireSize(options.WidthMm,options.HeightMm);
                transferred=roundTrip(accepted);
                final=validate(transferred);
                RequireSubmissionSelection(accepted,final);
                var geometry=measure(final);
                geometry.RequireSize(options.WidthMm,options.HeightMm);
                var result=final;final=null;
                return(result,geometry,string.Join(" ",attempts)+" Driver-validated custom settings survived the print-settings round trip.");
            }
            catch(Exception ex) when(ex is InvalidOperationException or System.ComponentModel.Win32Exception or InvalidDataException)
            {attempts.Add(variant+" not accepted: "+ex.Message);}
            finally
            {
                System.Security.Cryptography.CryptographicOperations.ZeroMemory(request);
                if(accepted is not null)System.Security.Cryptography.CryptographicOperations.ZeroMemory(accepted);
                if(transferred is not null)System.Security.Cryptography.CryptographicOperations.ZeroMemory(transferred);
                if(final is not null)System.Security.Cryptography.CryptographicOperations.ZeroMemory(final);
            }
        }
        throw new InvalidOperationException($"The driver did not accept the {options.WidthMm:0.##} x {options.HeightMm:0.##} mm envelope page. No print job was sent. Open Printer properties, select a custom original/output size matching the envelope, and keep the selected tray. "+string.Join(" ",attempts));
    }
    private static byte[] SubmissionMode(PrinterSettings settings,PageSettings page,EnvelopeOptions options)
    {
        byte[] requested=Mode(settings,page);
        if(!options.UsesWordEnvelopeCompatibility)return requested;
        try
        {
            // Reuse the accepted native custom form/private bytes; do not patch
            // it back to zero/Letter after the driver's validation.
            var accepted=ValidateMode(options.Printer,requested);
            ApplyMode(settings,page,accepted);return accepted;
        }
        finally{System.Security.Cryptography.CryptographicOperations.ZeroMemory(requested);}
    }
    private static void RequireSingleCopy(byte[] mode)
    {
        if(PrinterDevMode.CopiesValue(mode)!=1)throw new InvalidOperationException("Set driver copies to 1 in Printer properties. Use Copies per recipient in the app to control mailing quantities.");
    }
    private static string ReadMedia(string printer,byte[] mode)
    {
        try{return WithQueue(printer,q=>{using var converter=new PrintTicketConverter(q.FullName,PrintTicketConverter.MaxPrintSchemaVersion);return converter.ConvertDevModeToPrintTicket(mode,PrintTicketScope.JobScope).PageMediaType?.ToString()??"Driver default (type not reported)";});}
        catch{return "Driver settings (paper type not reported by Windows)";}
    }
    private static void RequireSubmissionSelection(byte[] expected,byte[] actual)
    {
        RequireSingleCopy(actual);
        if(PrinterDevMode.Tray(expected)!=PrinterDevMode.Tray(actual)||PrinterDevMode.DuplexValue(expected)!=PrinterDevMode.DuplexValue(actual))throw new InvalidOperationException("The driver changed the selected tray or sides setting before printing. Open Printer properties and try again.");
    }
    public Task<PreparedEnvelopeJob> PrepareAsync(EnvelopeSnapshot snapshot)
    {
        byte[]? selected=choices.Read(snapshot.Options);
        LastLayoutDiagnostics="Reading printer settings for "+snapshot.Options.Printer+".";
        return Sta.Run(()=>{try{return Prepare(snapshot,selected,text=>LastLayoutDiagnostics=text);}finally{if(selected is not null)System.Security.Cryptography.CryptographicOperations.ZeroMemory(selected);}});
    }
    private static PreparedEnvelopeJob Prepare(EnvelopeSnapshot snapshot,byte[]? selected,Action<string> diagnostics)
    {
        snapshot.Validate();var o=snapshot.Options;
        if(!o.IsLabel)throw new InvalidOperationException("Envelopes are printed from Word. Use Create Word document to print.");
        if(o.Printer.Contains("Print to PDF",StringComparison.OrdinalIgnoreCase)||o.Printer.Contains("XPS Document",StringComparison.OrdinalIgnoreCase)||o.Printer.Equals("Fax",StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("Choose a physical printer for direct mailing printing. Preview and history work without a printer.");
        var settings=Settings(o.Printer);var page=(PageSettings)settings.DefaultPageSettings.Clone();
        if(selected is not null)ApplyMode(settings,page,selected);else ConfigureInitialPage(settings,page,o);
        settings.PrintToFile=false;
        byte[] effective=Mode(settings,page);RequireSingleCopy(effective);
        string basePaper=PrinterDevMode.PaperDescription(effective);
        PrinterPageGeometry geometry;string reconciliation="";
        if(o.UsesWordEnvelopeCompatibility)
        {
            try
            {
                var resolved=ResolveEnvelopeMode(effective,o,mode=>ValidateMode(o.Printer,mode),mode=>
                {ApplyMode(settings,page,mode);return Mode(settings,page);},mode=>PrinterPageGeometry.Read(o.Printer,mode));
                System.Security.Cryptography.CryptographicOperations.ZeroMemory(effective);
                effective=resolved.Mode;geometry=resolved.Geometry;reconciliation=resolved.Diagnostics;
                ApplyMode(settings,page,effective);
            }
            catch(Exception ex){diagnostics(ex.Message);throw;}
        }
        else{geometry=PrinterPageGeometry.Read(o.Printer,effective);geometry.RequireSize(o.WidthMm,o.HeightMm);}
        if(selected is not null)RequireSubmissionSelection(selected,effective);
        var actual=snapshot.Copy();int? tray=PrinterDevMode.Tray(effective);
        actual.Options.TrayId=tray;actual.Options.TrayName=settings.PaperSources.Cast<PaperSource>().FirstOrDefault(t=>t.RawKind==tray)?.SourceName??"Printer default";
        string media=ReadMedia(o.Printer,effective);
        string measurements=reconciliation+" "+(o.UsesWordEnvelopeCompatibility?"Word-style compatibility: custom envelope document; original driver profile "+basePaper+"; document dimensions reconciled by the driver. ":"")+$"Settings source: {(selected is null?"printer defaults with document size and initial tray choice":"Printer properties for this app session")}. Requested layout: {o.WidthMm:0.##} x {o.HeightMm:0.##} mm. Accepted native paper: {PrinterDevMode.PaperDescription(effective)}. Accepted tray: {actual.Options.TrayName} (ID {tray}); media: {media}; duplex setting: {PrinterDevMode.DuplexValue(effective)}. "+geometry.Description;
        diagnostics(measurements);
        for(int index=0;index<snapshot.Pages;index++)
        {
            var drawing=MailingRenderer.RenderPage(snapshot,index,geometry.Imageable);
            diagnostics(measurements+" Page "+(index+1)+" address bounds: "+PrinterPageGeometry.Mm(drawing.ContentBounds)+" mm.");
            if(drawing.Warnings.Count>0)throw new InvalidOperationException("Page "+(index+1)+": "+string.Join(" ",drawing.Warnings)+"\n\nCheck Printer properties and the selected layout size. "+measurements);
        }
        return new(){Snapshot=actual,Imageable=geometry.Imageable,DriverState=effective,MediaType=media,SettingsDiagnostics=measurements};
    }
    internal static bool Fits(double pw,double ph,double w,double h)=>Math.Abs(pw-w)<0.8&&Math.Abs(ph-h)<0.8||Math.Abs(ph-w)<0.8&&Math.Abs(pw-h)<0.8;
    public Task<PrinterResult> SubmitAsync(PreparedEnvelopeJob job,string documentName,Action<PrinterProgress> progress,CancellationToken cancellation)=>Sta.Run(()=>
    {
        PropertyManager.App.PublicEdition.BlockLiveOperation();
        cancellation.ThrowIfCancellationRequested();var snapshot=job.Snapshot;var o=snapshot.Options;
        if(!o.IsLabel)throw new InvalidOperationException("Envelopes are printed from Word. Use Create Word document to print.");
        var settings=Settings(o.Printer);var page=(PageSettings)settings.DefaultPageSettings.Clone();
        ApplyMode(settings,page,(byte[])job.DriverState!);settings.Copies=1;settings.PrintToFile=false;
        byte[] submission=SubmissionMode(settings,page,o);RequireSubmissionSelection((byte[])job.DriverState!,submission);
        var finalGeometry=PrinterPageGeometry.Read(o.Printer,submission);finalGeometry.RequireSize(o.WidthMm,o.HeightMm);
        if(finalGeometry.Imageable!=job.Imageable)throw new InvalidOperationException("The printable area changed before submission. Check printer settings again. No print job was sent.");
        string submissionMedia=ReadMedia(o.Printer,submission);
        progress(new(0,snapshot.Pages,null,job.SettingsDiagnostics+" Final submission settings rechecked: "+PrinterDevMode.PaperDescription(submission)+"; "+submissionMedia+"; using the selected driver settings."));
        using var document=new PrintDocument{PrinterSettings=settings,DefaultPageSettings=page,DocumentName=documentName,OriginAtMargins=false,PrintController=new StandardPrintController()};
        int pages=0;int? jobId=null;
        document.QueryPageSettings+=(_,e)=>
        {
            byte[]? pageMode=null;
            try
            {
                // Check the page settings supplied to the print controller as well.
                pageMode=SubmissionMode(settings,e.PageSettings,o);
                RequireSubmissionSelection(submission,pageMode);
                var pageGeometry=PrinterPageGeometry.Read(o.Printer,pageMode);pageGeometry.RequireSize(o.WidthMm,o.HeightMm);
                if(pageGeometry.Imageable!=job.Imageable)throw new InvalidOperationException("The page printable area changed. Check Printer properties before retrying.");
                if(pages==0)progress(new(0,snapshot.Pages,null,"First-page tray and sides settings checked. Printer-side loaded stock must match the driver settings."));
            }
            catch(Exception ex){throw new InvalidOperationException("Printer settings changed while starting the page. Printing was stopped before rendering this page; check the queue before retrying.",ex);}
            finally{if(pageMode is not null)System.Security.Cryptography.CryptographicOperations.ZeroMemory(pageMode);}
        };
        document.PrintPage+=(_,e)=>
        {
            if(cancellation.IsCancellationRequested){e.Cancel=true;return;}
            jobId??=FindJob(o.Printer,documentName);
            progress(new(pages,snapshot.Pages,jobId));
            var drawing=MailingRenderer.RenderPage(snapshot,pages,job.Imageable);
            if(drawing.Warnings.Count>0)throw new InvalidOperationException("The mailing layout no longer fits the validated page.");
            using var stream=new MemoryStream(EnvelopeRenderer.Png(drawing));using var bitmap=System.Drawing.Image.FromStream(stream);
            var graphics=e.Graphics??throw new InvalidOperationException("The printer did not provide a drawing surface.");
            graphics.PageUnit=System.Drawing.GraphicsUnit.Inch;
            graphics.DrawImage(bitmap,new System.Drawing.RectangleF((float)(-job.Imageable.X/96),(float)(-job.Imageable.Y/96),(float)(o.WidthMm/25.4),(float)(o.HeightMm/25.4)));
            pages++;progress(new(pages,snapshot.Pages,jobId));e.HasMorePages=pages<snapshot.Pages;
        };
        document.Print();
        jobId??=FindJob(o.Printer,documentName);
        return new PrinterResult(pages,jobId,cancellation.IsCancellationRequested||pages<snapshot.Pages);
    });
    private static int? FindJob(string name,string document)
    {
        try{return WithQueue<int?>(name,q=>q.GetPrintJobInfoCollection().FirstOrDefault(j=>j.Name==document)?.JobIdentifier);}catch{return null;}
    }
    public Task<string> QueueStatusAsync(string printer,int jobId,string documentName)=>Task.Run(()=>WithQueue(printer,q=>
    {
        PrintSystemJobInfo? job=q.GetPrintJobInfoCollection().FirstOrDefault(j=>j.JobIdentifier==jobId&&j.Name==documentName);
        if(job is null)return "Not in queue — physical outcome unconfirmed";
        job.Refresh();return DescribeStatus(job.JobStatus);
    }));
    internal static string DescribeStatus(PrintJobStatus status)
    {
        if(status.HasFlag(PrintJobStatus.Printed))return "Printer-reported printed";
        if(status.HasFlag(PrintJobStatus.Completed))return "Delivered to printer — physical outcome unconfirmed";
        if(status.HasFlag(PrintJobStatus.Error)||status.HasFlag(PrintJobStatus.PaperOut)||status.HasFlag(PrintJobStatus.Offline)||status.HasFlag(PrintJobStatus.UserIntervention))return "Queue needs attention: "+status;
        if(status.HasFlag(PrintJobStatus.Deleted)||status.HasFlag(PrintJobStatus.Deleting))return "Queue reports deleted — may be partially printed";
        return "Queue reports: "+status;
    }
    private static T WithQueue<T>(string name,Func<PrintQueue,T> action)
    {
        if(name.StartsWith(@"\\"))
        {int split=name.LastIndexOf('\\');using var server=new PrintServer(name[..split]);using var queue=server.GetPrintQueue(name[(split+1)..]);return action(queue);}
        using(var server=new LocalPrintServer())using(var queue=server.GetPrintQueue(name))return action(queue);
    }
    private static byte[] Mode(PrinterSettings settings,PageSettings page)
    {
        IntPtr handle=settings.GetHdevmode(page);
        if(handle==IntPtr.Zero)throw new InvalidOperationException("The printer did not return native settings.");
        try{return ReadHdevmode(handle);}finally{FreeHdevmode(handle);}
    }
    internal static byte[] ReadHdevmode(IntPtr handle)
    {
        IntPtr pointer=GlobalLock(handle);
        if(pointer==IntPtr.Zero)throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        try{return Bytes(pointer,checked((int)GlobalSize(handle).ToUInt64()));}finally{GlobalUnlock(handle);}
    }
    internal static IntPtr CreateHdevmode(byte[] bytes)
    {
        IntPtr handle=GlobalAlloc(0x2,(UIntPtr)(uint)bytes.Length);
        if(handle==IntPtr.Zero)throw new OutOfMemoryException("Could not allocate printer settings.");
        IntPtr pointer=GlobalLock(handle);
        if(pointer==IntPtr.Zero){int error=Marshal.GetLastWin32Error();GlobalFree(handle);throw new System.ComponentModel.Win32Exception(error);}
        try{Marshal.Copy(bytes,0,pointer,bytes.Length);}catch{GlobalUnlock(handle);GlobalFree(handle);throw;}
        GlobalUnlock(handle);return handle;
    }
    internal static void FreeHdevmode(IntPtr handle)=>GlobalFree(handle);
    private static byte[] Bytes(IntPtr pointer,int capacity)
    {
        if(capacity<96)throw new InvalidDataException("The printer returned incomplete native settings.");
        int size=Marshal.ReadInt16(pointer,68)&0xffff,extra=Marshal.ReadInt16(pointer,70)&0xffff;
        if(size<96||size+extra>capacity)throw new InvalidDataException("The printer returned invalid native settings lengths.");
        var bytes=new byte[checked(size+extra)];Marshal.Copy(pointer,bytes,0,bytes.Length);return bytes;
    }
    private static void ApplyMode(PrinterSettings settings,PageSettings page,byte[] bytes)
    {
        // Get/SetHdevmode exchange HGLOBAL handles, not raw pointers.
        IntPtr handle=CreateHdevmode(bytes);
        try{settings.SetHdevmode(handle);page.SetHdevmode(handle);}finally{FreeHdevmode(handle);}
    }
    private static byte[] ValidateMode(string name,byte[] bytes)
    {
        if(!OpenPrinter(name,out var handle,IntPtr.Zero))throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        IntPtr input=Marshal.AllocHGlobal(bytes.Length),output=IntPtr.Zero;
        try
        {
            Marshal.Copy(bytes,0,input,bytes.Length);int size=DocumentProperties(IntPtr.Zero,handle,name,IntPtr.Zero,IntPtr.Zero,0);
            if(size<=0)throw new InvalidOperationException("The printer driver did not report a valid settings size.");
            output=Marshal.AllocHGlobal(size);
            if(DocumentProperties(IntPtr.Zero,handle,name,output,input,0x2|0x8)!=1)throw new InvalidOperationException("The printer rejected these envelope settings.");
            return Bytes(output,size);
        }finally{Marshal.FreeHGlobal(input);if(output!=IntPtr.Zero)Marshal.FreeHGlobal(output);ClosePrinter(handle);}
    }
    [DllImport("kernel32.dll",SetLastError=true)] private static extern IntPtr GlobalAlloc(uint flags,UIntPtr bytes);
    [DllImport("kernel32.dll",SetLastError=true)] private static extern IntPtr GlobalLock(IntPtr handle);
    [DllImport("kernel32.dll",SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] private static extern bool GlobalUnlock(IntPtr handle);
    [DllImport("kernel32.dll",SetLastError=true)] private static extern IntPtr GlobalFree(IntPtr handle);
    [DllImport("kernel32.dll",SetLastError=true)] private static extern UIntPtr GlobalSize(IntPtr handle);
    [DllImport("winspool.drv",CharSet=CharSet.Unicode,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] private static extern bool OpenPrinter(string name,out IntPtr handle,IntPtr defaults);
    [DllImport("winspool.drv",SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] private static extern bool ClosePrinter(IntPtr handle);
    [DllImport("winspool.drv",CharSet=CharSet.Unicode)] private static extern int DocumentProperties(IntPtr owner,IntPtr printer,string name,IntPtr output,IntPtr input,int mode);
}
internal static class Sta
{
    internal static Task<T> Run<T>(Func<T> action)
    {
        var completion=new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread=new Thread(()=>{try{completion.SetResult(action());}catch(Exception ex){completion.SetException(ex);}}){IsBackground=true};
        thread.SetApartmentState(ApartmentState.STA);thread.Start();return completion.Task;
    }
}
