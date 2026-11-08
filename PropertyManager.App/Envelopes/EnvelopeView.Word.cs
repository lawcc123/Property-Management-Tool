using System.Diagnostics;
using System.IO;
using System.Windows;
using Microsoft.Win32;
namespace PropertyManager.App.Envelopes;
public partial class EnvelopeView
{
    private async void CreateWord_Click(object sender,RoutedEventArgs e)
    {
        if(busy)return;
        try
        {
            Snapshot().Validate();
            var dialog=new SaveFileDialog{Title="Save envelopes for printing in Word",Filter="Word document (*.docx)|*.docx",DefaultExt=".docx",FileName="PropertyManager-envelopes-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N")+".docx"};
            if(dialog.ShowDialog(OwnerWindow)!=true)return;
            await CreateWordDocumentAsync(dialog.FileName,true);
        }
        catch(Exception ex){Error(ex);}
    }
    internal async Task CreateWordDocumentAsync(string path,bool openInWord=false)
    {
        if(busy)throw new InvalidOperationException("Finish the current operation first.");
        var snapshot=Snapshot();if(snapshot.Options.IsLabel)throw new InvalidOperationException("Choose Envelopes to create a Word document.");snapshot.Validate();
        string? temp=null;EnvelopeHistoryEntry? entry=null;bool created=false;
        try
        {
            SetBusy(true);cancellation=new();var token=cancellation.Token;
            PrintProgress.IsIndeterminate=true;StatusText.Text="Creating the envelope document for Word.";
            entry=store.Create(snapshot,"Word document",parentHistory);store.SaveOptions(snapshot.Options);
            temp=path+"."+Guid.NewGuid().ToString("N")+".tmp";string staged=temp;
            await Sta.Run(()=>{EnvelopeWordDocument.Write(staged,snapshot,(page,total)=>Dispatcher.BeginInvoke(()=>{PrintProgress.IsIndeterminate=false;PrintProgress.Value=100.0*page/total;}),token);return true;});
            token.ThrowIfCancellationRequested();File.Move(temp,path,true);temp=null;created=true;
            store.Update(entry.Id,"Word document created",$"Saved {snapshot.Pages} envelope page(s): {path}. Open in Word, select the printer and Bypass Tray in Printer Properties, then print.");
            StatusText.Text="Envelope document saved. Print it from Word and select Bypass Tray in Word's Printer Properties.";
            if(openInWord)Process.Start(new ProcessStartInfo(path){UseShellExecute=true});
        }
        catch(Exception ex)
        {
            if(entry is not null&&!created)try{store.Update(entry.Id,"Word generation stopped",ex.Message);}catch{}
            StatusText.Text=created?"The Word document was saved. Open it from the saved location: "+path:"Word document: "+ex.Message;
            throw new InvalidOperationException(StatusText.Text,ex);
        }
        finally
        {
            if(temp is not null&&File.Exists(temp))try{File.Delete(temp);}catch(IOException){}catch(UnauthorizedAccessException){}
            PrintProgress.IsIndeterminate=false;cancellation?.Dispose();cancellation=null;SetBusy(false);RefreshHistory();
        }
    }
}
