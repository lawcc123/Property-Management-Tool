using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace PropertyManager.App;

public partial class MainWindow
{
    private void UpdateBackupFolder()
    {
        RegistryBackupFolderPath.Text=settings.EffectiveRegistryBackupFolder();
        RegistryBackupFolderPath.ToolTip=RegistryBackupFolderPath.Text;
        DefaultBackupFolderButton.IsEnabled=!running&&!string.IsNullOrWhiteSpace(settings.RegistryBackupFolder);
    }
    internal void SetRegistryBackupFolder(string? folder)
    {
        if(running)throw new InvalidOperationException("Finish or cancel the sync before changing the backup folder.");
        string value=string.IsNullOrWhiteSpace(folder)?"":Path.GetFullPath(folder);
        if(!string.IsNullOrWhiteSpace(folder)&&!Path.IsPathFullyQualified(folder))throw new InvalidOperationException("Choose a full registry backup folder path.");
        if(File.Exists(value))throw new InvalidOperationException("Choose a folder rather than a file for registry backups.");
        string previous=settings.RegistryBackupFolder;
        try{settings.RegistryBackupFolder=value;settings.Save();}
        catch{settings.RegistryBackupFolder=previous;throw;}
        UpdateBackupFolder();
        StageText.Text=value.Length==0?"Backup location reset to Registry Backups beside the workbook.":"Registry backups will be saved in the selected folder. Existing backups have not been moved.";
    }
    private void ChangeBackupFolder_Click(object sender,RoutedEventArgs e)
    {
        if(running)return;
        var picker=new OpenFolderDialog{Title="Choose the registry backup folder",Multiselect=false};
        string initial=settings.EffectiveRegistryBackupFolder();if(Directory.Exists(initial))picker.InitialDirectory=initial;
        if(picker.ShowDialog(this)!=true)return;
        try{SetRegistryBackupFolder(picker.FolderName);}catch(Exception ex){MessageBox.Show(this,ex.Message,"Registry backup folder",MessageBoxButton.OK,MessageBoxImage.Warning);}
    }
    private void DefaultBackupFolder_Click(object sender,RoutedEventArgs e)
    {
        try{SetRegistryBackupFolder(null);}catch(Exception ex){MessageBox.Show(this,ex.Message,"Registry backup folder",MessageBoxButton.OK,MessageBoxImage.Warning);}
    }
}
