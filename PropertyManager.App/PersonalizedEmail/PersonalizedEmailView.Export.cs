using System.Diagnostics;
using System.IO;
using System.Windows;
using Microsoft.Win32;
namespace PropertyManager.App.PersonalizedEmail;
public partial class PersonalizedEmailView
{
    private string exportFolder="";
    private void RefreshExportFolder()
    {
        ExportFolderText.Text=exportFolder.Length==0?"Save personalized Word/PDF files to a folder you choose. Each dated export has one folder per personalized template, containing all included units (including print-only records). Shared files are excluded.": "Last export folder: "+exportFolder;
        OpenExportFolderButton.IsEnabled=!deliveryBusy&&exportFolder.Length>0;
        SaveAttachmentsButton.IsEnabled=!deliveryBusy&&documents.Any(d=>d.IsWord&&d.Personalize)&&recipients.Any(r=>r.Included);
    }
    private async void SaveAttachments_Click(object s,RoutedEventArgs e)
    {
        if(deliveryBusy)return;
        try
        {
            var draft=ReadDraft();if(!draft.Documents.Any(d=>d.IsWord&&d.Personalize))throw new InvalidOperationException("Add a Word template with Personalize checked first. Shared files are not exported.");
            var picker=new OpenFolderDialog{Title="Choose where to save generated attachments"};if(picker.ShowDialog(OwnerWindow)!=true)return;
            DeliveryBusy(true);ExportProgress.Visibility=Visibility.Visible;ExportProgress.Value=0;
            var progress=new Progress<(string Text,int Done,int Total)>(p=>{StatusText.Text=p.Text;ExportProgress.Value=p.Total==0?0:100.0*p.Done/p.Total;});
            var token=deliveryStop!.Token;
            var result=await Envelopes.Sta.Run(()=>new CampaignAttachmentExport(()=>new WordCampaignDocuments()).Generate(draft,picker.FolderName,(m,d,t)=>((IProgress<(string Text,int Done,int Total)>)progress).Report((m,d,t)),token));
            exportFolder=result.Folder;RefreshExportFolder();
            StatusText.Text=result.Complete?$"Saved {result.Files} file(s) for {result.Recipients} recipient set(s). Use Open folder to inspect, copy or delete them. No emails sent.":result.Problem+" Use Open folder to inspect or delete partial files.";
            if(result.Complete)MessageBox.Show(OwnerWindow,StatusText.Text+"\n\n"+result.Folder,"Attachments saved",MessageBoxButton.OK,MessageBoxImage.Information);
            else if(!token.IsCancellationRequested)Warn(StatusText.Text);
        }
        catch(OperationCanceledException){StatusText.Text="Attachment export cancelled. No emails sent.";}
        catch(Exception ex){Error(ex);}
        finally{ExportProgress.Visibility=Visibility.Collapsed;if(deliveryBusy)DeliveryBusy(false);RefreshExportFolder();}
    }
    private void OpenExportFolder_Click(object s,RoutedEventArgs e)
    {
        try{if(exportFolder.Length==0||!Directory.Exists(exportFolder))throw new DirectoryNotFoundException("This export folder is unavailable. Check the drive connection or save the attachments again.");Process.Start(new ProcessStartInfo(exportFolder){UseShellExecute=true});}catch(Exception ex){Error(ex);}
    }
}
