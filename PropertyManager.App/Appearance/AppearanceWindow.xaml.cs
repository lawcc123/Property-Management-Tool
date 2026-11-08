using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace PropertyManager.App.Appearance;
public partial class AppearanceWindow : Window
{
    private readonly Action<AppearanceSettings> preview;
    private readonly ObservableCollection<WorkspaceItem> items=[];
    private bool loading;
    internal AppearanceSettings? SavedSettings {get;private set;}
    internal AppearanceWindow(AppearanceSettings current,Action<AppearanceSettings> preview)
    {
        this.preview=preview;InitializeComponent();Height=Math.Min(760,SystemParameters.WorkArea.Height-50);PresetCombo.ItemsSource=AppearanceSettings.Presets;Populate(current);
    }
    private void Populate(AppearanceSettings settings)
    {
        loading=true;AccentText.Text=settings.Accent;SidebarText.Text=settings.Sidebar;
        PresetCombo.SelectedItem=AppearanceSettings.Presets.FirstOrDefault(p=>p.Accent==settings.Accent&&p.Sidebar==settings.Sidebar);
        items.Clear();foreach(string id in settings.WorkspaceOrder)items.Add(AppearanceSettings.Workspaces.Single(w=>w.Id==id));WorkspaceList.ItemsSource=items;WorkspaceList.SelectedIndex=0;loading=false;UpdateButtons();Swatches(settings);
    }
    private void Swatches(AppearanceSettings settings){AccentSwatch.Background=new SolidColorBrush(AppearanceSettings.Parse(settings.Accent));SidebarSwatch.Background=new SolidColorBrush(AppearanceSettings.Parse(settings.Sidebar));}
    private AppearanceSettings Read(){var s=new AppearanceSettings{Accent=AccentText.Text.Trim(),Sidebar=SidebarText.Text.Trim(),WorkspaceOrder=items.Select(w=>w.Id).ToList()};s.Validate();return s;}
    private bool Preview()
    {
        try{var settings=Read();preview(settings);Swatches(settings);ErrorText.Text="";return true;}catch(Exception ex){ErrorText.Text=ex.Message;return false;}
    }
    private void Preset_Changed(object s,SelectionChangedEventArgs e){if(!loading&&PresetCombo.SelectedItem is ThemePreset preset){AccentText.Text=preset.Accent;SidebarText.Text=preset.Sidebar;Preview();}}
    private void Preview_Click(object s,RoutedEventArgs e)=>Preview();
    private void IgnoreWheel(object s,MouseWheelEventArgs e){if(s is ComboBox{IsDropDownOpen:false})e.Handled=true;}
    private void Workspace_Changed(object s,SelectionChangedEventArgs e){if(IsInitialized)UpdateButtons();}
    private void UpdateButtons(){if(MoveUpButton is null)return;MoveUpButton.IsEnabled=WorkspaceList.SelectedIndex>0;MoveDownButton.IsEnabled=WorkspaceList.SelectedIndex>=0&&WorkspaceList.SelectedIndex<items.Count-1;}
    internal void Move(int direction)
    {
        int old=WorkspaceList.SelectedIndex,next=old+direction;if(old<0||next<0||next>=items.Count)return;
        items.Move(old,next);WorkspaceList.SelectedIndex=next;UpdateButtons();Preview();
    }
    private void MoveUp_Click(object s,RoutedEventArgs e)=>Move(-1);
    private void MoveDown_Click(object s,RoutedEventArgs e)=>Move(1);
    private void Defaults_Click(object s,RoutedEventArgs e){Populate(new());Preview();}
    private void Save_Click(object s,RoutedEventArgs e)
    {
        try{var settings=Read();settings.Save();preview(settings);SavedSettings=settings;DialogResult=true;}catch(Exception ex){ErrorText.Text=ex.Message;}
    }
}
