using System.Windows;
using PropertyManager.App.Settings;
namespace PropertyManager.App;
public partial class MainWindow
{
    private bool SharedSettingsBusy()=>CorporationSettingsBusy||envelopeView?.IsBusy==true||personalizedEmailView?.IsBusy==true;
    private void Settings_Click(object sender,RoutedEventArgs e)
    {
        if(SharedSettingsBusy()){MessageBox.Show(this,"Finish the current operation before storage maintenance.","Settings");return;}
        var editor=new CentralSettingsWindow(SharedSettingsBusy){Owner=this};
        editor.ShowDialog();buildingLinkView?.ReloadSharedFolders();emailListView?.ReloadSavedFolders();ownerMailingView?.ReloadSavedFolders();outlookSyncView?.ReloadSavedFolders();operationsView?.ReloadEmailSettings();
    }
}
