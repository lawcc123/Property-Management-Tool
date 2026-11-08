using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using PropertyManager.App.Services;
namespace PropertyManager.App.Settings;

internal sealed class CentralSettingsWindow:Window
{
    internal readonly ScrollViewer StorageScroll=new();
    internal readonly TextBlock StorageStatus=new(){TextWrapping=TextWrapping.Wrap};
    internal readonly StackPanel StoragePanel=new(){Margin=new Thickness(20)};
    internal readonly Button MigrateButton;
    private readonly TextBlock result=new(){TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,12,0,0)};
    private readonly Func<bool> busy;
    internal CentralSettingsWindow(Func<bool> busy)
    {
        this.busy=busy;
        Title="Storage & backup - PropertyManagerDemo";Width=940;Height=710;MinWidth=720;MinHeight=500;WindowStartupLocation=WindowStartupLocation.CenterOwner;
        SetResourceReference(BackgroundProperty,"ThemePage");Foreground=new SolidColorBrush(Color.FromRgb(23,43,58));
        var root=new DockPanel{Margin=new Thickness(24)};Content=root;
        var heading=new StackPanel{Margin=new Thickness(0,0,0,18)};
        heading.Children.Add(new TextBlock{Text="Storage & backup",FontSize=28,FontWeight=FontWeights.SemiBold});
        heading.Children.Add(Caption("Back up this computer's app data to a folder you choose, and verify stored settings."));
        DockPanel.SetDock(heading,Dock.Top);root.Children.Add(heading);
        var close=new Button{Content="Done",IsCancel=true,HorizontalAlignment=HorizontalAlignment.Right,Margin=new Thickness(0,16,0,0)};
        close.Click+=(_,_)=>Close();DockPanel.SetDock(close,Dock.Bottom);root.Children.Add(close);
        var card=new Border{Child=StoragePanel,Padding=new Thickness(0)};card.SetResourceReference(StyleProperty,"Card");
        StorageScroll.Content=card;StorageScroll.VerticalScrollBarVisibility=ScrollBarVisibility.Auto;StorageScroll.HorizontalScrollBarVisibility=ScrollBarVisibility.Disabled;StorageScroll.Padding=new Thickness(0,0,12,0);root.Children.Add(StorageScroll);
        StoragePanel.Children.Add(new TextBlock{Text="Storage on this computer",FontSize=19,FontWeight=FontWeights.SemiBold});
        StoragePanel.Children.Add(Caption(AppSettings.DataDirectory));StoragePanel.Children.Add(StorageStatus);
        ButtonAction("Back up app data",async()=>await Backup());
        ButtonAction("Verify stored settings",()=>StorageStatus.Text=SettingsDatabase.Verify());
        MigrateButton=ButtonAction("Migrate settings and verify",async()=>await Migrate());
        StoragePanel.Children.Add(Caption("Migration moves shared settings and the contractor address book into app.db. Histories, drafts, signatures, passwords and browser sessions keep their existing storage. Migration uses the same app-data backup method and saves its verification report in your selected backup folder."));
        ButtonAction("Open app-data folder",()=>{Directory.CreateDirectory(AppSettings.DataDirectory);System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(AppSettings.DataDirectory){UseShellExecute=true});});
        StoragePanel.Children.Add(Caption("Backups require the original Windows account and computer. Passwords, browser sessions, external workbooks, exported documents and Windows scheduled tasks are not included. Existing backups are retained. This page does not clear history or restore older data."));
        StoragePanel.Children.Add(result);RefreshStorage();
    }
    private static TextBlock Caption(string text)=>new(){Text=text,TextWrapping=TextWrapping.Wrap,Foreground=new SolidColorBrush(Color.FromRgb(102,121,133)),Margin=new Thickness(0,8,0,14)};
    private Button ButtonAction(string label,Action action)
    {
        var button=new Button{Content=label,HorizontalAlignment=HorizontalAlignment.Left,Margin=new Thickness(0,0,0,10),Padding=new Thickness(14,9,14,9)};
        button.Click+=(_,_)=>{try{if(busy())throw new InvalidOperationException("Finish the current operation before storage maintenance.");action();}catch(Exception ex){MessageBox.Show(this,ex.Message,"Storage & backup",MessageBoxButton.OK,MessageBoxImage.Warning);}};StoragePanel.Children.Add(button);return button;
    }
    private void RefreshStorage()
    {
        try{bool active=SettingsDatabase.Active;MigrateButton.IsEnabled=!active;MigrateButton.ToolTip=active?"Shared settings have already been migrated. Use Verify stored settings to check them.":null;StorageStatus.Text=active?SettingsDatabase.Verify():"Existing file storage is active. You can back it up now, or migrate shared settings when ready.";}
        catch(Exception ex){MigrateButton.IsEnabled=false;StorageStatus.Text="Storage needs attention: "+ex.Message;}
    }
    private string? ChooseBackupFolder()
    {
        var picker=new OpenFolderDialog{Title="Choose a folder for the app-data backup"};return picker.ShowDialog(this)==true?picker.FolderName:null;
    }
    private async Task Migrate()
    {
        if(busy())return;
        string? folder=ChooseBackupFolder();if(folder is null)return;
        if(MessageBox.Show(this,"Create a verified app-data backup in the selected folder, then migrate shared settings?\n\nHistories and drafts stay in place. No Outlook changes or emails are sent.","Migrate shared settings",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return;
        StoragePanel.IsEnabled=false;Closing+=PreventClose;
        try{result.Text="Backing up, importing and comparing settings...";result.Text=await Task.Run(()=>SettingsDatabase.Migrate(folder));RefreshStorage();}
        catch(Exception ex){result.Text="Migration did not complete: "+ex.Message+" Original files were retained.";RefreshStorage();}
        finally{Closing-=PreventClose;StoragePanel.IsEnabled=true;}
    }
    private async Task Backup()
    {
        string? folder=ChooseBackupFolder();if(folder is null)return;
        StoragePanel.IsEnabled=false;Closing+=PreventClose;
        try{result.Text="Backing up settings, drafts and histories...";string path=await Task.Run(()=>AppDataBackup.Create(folder));result.Text="Backup complete: "+path+"\nKeep it for this Windows account and computer.";}
        catch(Exception ex){result.Text="Backup did not complete: "+ex.Message;}
        finally{Closing-=PreventClose;StoragePanel.IsEnabled=true;}
    }
    private void PreventClose(object? sender,System.ComponentModel.CancelEventArgs e){e.Cancel=true;}
}
