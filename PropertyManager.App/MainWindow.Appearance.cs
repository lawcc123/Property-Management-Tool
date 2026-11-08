using System.Windows;
using PropertyManager.App.Appearance;

namespace PropertyManager.App;
public partial class MainWindow
{
    private AppearanceSettings appearance=new();
    private int activeWorkspace;
    private string appearanceLoadError="";
    private void InitializeAppearance()
    {
        try{appearance=AppearanceSettings.Load();}catch(Exception ex){appearanceLoadError="Saved appearance could not be loaded. Defaults are being used. "+ex.Message;SettingsNavigationButton.ToolTip=appearanceLoadError;}
        ApplyAppearance(appearance);
    }
    internal void ApplyAppearance(AppearanceSettings settings)
    {
        settings.Apply();var entries=new Dictionary<string,(FrameworkElement Item,System.Windows.Controls.Button Button,int Page)>
        {
            ["registry"]=(RegistryNavigationItem,RegistryNavigation,0),["email"]=(EmailListNavigation,EmailListNavigation,1),["owners"]=(OwnerMailingNavigation,OwnerMailingNavigation,4),["buildinglink"]=(BuildingLinkNavigation,BuildingLinkNavigation,2),["printing"]=(EnvelopeNavigation,EnvelopeNavigation,3),["outlook"]=(OutlookSyncNavigation,OutlookSyncNavigation,5),["reminders"]=(OperationsNavigation,OperationsNavigation,6),["personalized"]=(PersonalizedEmailNavigation,PersonalizedEmailNavigation,7)
        };
        WorkspaceItems.Children.Clear();int index=0;
        foreach(string id in settings.WorkspaceOrder)
        {
            var entry=entries[id];entry.Item.Margin=new Thickness(0,index++==0?0:5,0,0);WorkspaceItems.Children.Add(entry.Item);entry.Button.Background=PageBrush(entry.Page==activeWorkspace);
        }
    }
    private void Appearance_Click(object sender,RoutedEventArgs e)
    {
        if(SharedSettingsBusy()){MessageBox.Show(this,"Finish the current operation before changing appearance.","Appearance & workspace");return;}
        try{EditAppearance(this);}catch(Exception ex){MessageBox.Show(this,ex.Message,"Appearance & workspace");}
    }
    private void EditAppearance(Window owner)
    {
        var original=appearance.Copy();var editor=new AppearanceWindow(original,ApplyAppearance){Owner=owner};
        if(appearanceLoadError.Length>0)((System.Windows.Controls.TextBlock)editor.FindName("ErrorText")).Text=appearanceLoadError;
        try
        {
            if(editor.ShowDialog()==true&&editor.SavedSettings is {} saved){appearance=saved;appearanceLoadError="";SettingsNavigationButton.ToolTip=null;}
            else ApplyAppearance(original);
        }
        catch{ApplyAppearance(original);throw;}
    }
}
