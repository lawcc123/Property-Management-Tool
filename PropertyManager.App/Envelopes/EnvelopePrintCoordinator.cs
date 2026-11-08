namespace PropertyManager.App.Envelopes;

internal sealed class EnvelopePrintCoordinator(EnvelopeStore store,IEnvelopePrinter printer)
{
    internal async Task<EnvelopeHistoryEntry> PrintAsync(EnvelopeSnapshot snapshot,string kind,string? parent,
        Func<PreparedEnvelopeJob,Task<bool>> confirm,Action<PrinterProgress> progress,CancellationToken cancellation,Func<PrinterMediaSetupException,Task<bool>>? configureMedia=null)
    {
        // Persist the immutable request before validating or touching the spooler.
        var entry=store.Create(snapshot,kind,parent);bool submitting=false;
        try
        {
            using var prepared=await PrepareWithMediaSetupAsync(printer,entry.Snapshot,configureMedia,cancellation);
            store.Update(entry.Id,"Validated","Printer, tray, size, printable area and mailing media checked. "+prepared.MediaType+" "+prepared.SettingsDiagnostics);
            if(!await confirm(prepared)){return store.Update(entry.Id,"Cancelled before submission","User cancelled the confirmation. No print job was submitted.");}
            cancellation.ThrowIfCancellationRequested();
            store.Update(entry.Id,"Submitting","Sending mailing pages to the Windows print queue. Physical printing is not confirmed.");submitting=true;
            int? recordedJob=null;
            var result=await printer.SubmitAsync(prepared,entry.DocumentName,p=>
            {
                if(p.Settings is not null)store.Update(entry.Id,"Driver settings verified",p.Settings);
                if(p.JobId is int id&&recordedJob!=id){store.Update(entry.Id,"Spooling","Windows job ID recorded; physical printing is not confirmed.",id);recordedJob=id;}
                progress(p);
            },cancellation);
            return store.Update(entry.Id,result.Cancelled?"Cancellation requested — may be partial":"Submitted to printer",
                $"{result.PagesRendered} of {entry.Pages} mailing pages rendered to the spooler. "+(result.Cancelled?"Check the printer before reprinting.":"Physical printing is not confirmed."),result.JobId);
        }
        catch(OperationCanceledException)
        {return store.Update(entry.Id,submitting?"Cancellation requested — may be partial":"Cancelled before submission",submitting?"Check the printer before reprinting.":"No job submitted.");}
        catch(Exception ex)
        {store.Update(entry.Id,submitting?"Submission failed — may be partial":"Validation failed",ex.Message);throw;}
    }
    internal static async Task<PreparedEnvelopeJob> PrepareWithMediaSetupAsync(IEnvelopePrinter printer,EnvelopeSnapshot snapshot,
        Func<PrinterMediaSetupException,Task<bool>>? configureMedia,CancellationToken cancellation)
    {
        cancellation.ThrowIfCancellationRequested();
        try{return await printer.PrepareAsync(snapshot);}
        catch(PrinterMediaSetupException ex) when(configureMedia is not null)
        {
            // Retry preparation only, once, before confirmation or any spooler submission.
            cancellation.ThrowIfCancellationRequested();
            if(!await configureMedia(ex))throw new OperationCanceledException("Printer media setup was cancelled.");
            cancellation.ThrowIfCancellationRequested();
            return await printer.PrepareAsync(snapshot);
        }
    }

}
