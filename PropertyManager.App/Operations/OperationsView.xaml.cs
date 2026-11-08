using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using PropertyManager.App.Services;

namespace PropertyManager.App.Operations;

public partial class OperationsView : UserControl
{
    private readonly OperationsRunner runner=new();
    private readonly Action manageCredentials;
    private OperationsSettings settings;
    private string workbookPath="";
    private JsonElement? prepared;
    private bool loading=true,scheduleBusy,processing;
    internal bool IsBusy=>processing||runner.IsRunning||scheduleBusy;
    internal void Cancel()=>runner.Cancel();
    private Window OwnerWindow=>Window.GetWindow(this);
    internal OperationsView(Action manageCredentials)
    {
        this.manageCredentials=manageCredentials;InitializeComponent();foreach(var day in Enum.GetValues<DayOfWeek>())ScheduleDays.Children.Add(new CheckBox{Content=day.ToString()[..3],Tag=(int)day,Margin=new Thickness(0,0,12,0)});settings=OperationsSettings.Load();Populate();RefreshCredentials();RefreshHistory();
    }
    private void Populate()
    {
        loading=true;SelectWorkbook(settings.Workbook);DaysText.Text=settings.DaysAhead.ToString();PartialCheck.IsChecked=settings.AllowPartial;
        TimeText.Text=settings.Time;foreach(var day in ScheduleDays.Children.OfType<CheckBox>())day.IsChecked=settings.Days.Contains((int)day.Tag);loading=false;
    }
    private OperationsSettings Read()
    {
        if(!int.TryParse(DaysText.Text,out int days))throw new InvalidDataException("Days ahead must be a number.");
        return new(){Workbook=workbookPath,PropertyName=OperationsSettings.SharedProperty,DaysAhead=days,AllowPartial=PartialCheck.IsChecked==true,SmtpServer=settings.SmtpServer,SmtpPort=settings.SmtpPort,From=settings.From,Recipients=[..settings.Recipients],Cc=[..settings.Cc],Subject=settings.Subject,Time=TimeText.Text.Trim(),Days=ScheduleDays.Children.OfType<CheckBox>().Where(c=>c.IsChecked==true).Select(c=>(int)c.Tag).ToList()};
    }
    private void InvalidatePreview(){prepared=null;if(SendButton is not null){SendButton.IsEnabled=EmailPreviewButton.IsEnabled=false;SummaryText.Text="Settings changed. Refresh to create a new preview.";WarningText.Text="";}}
    private void Changed(object sender,RoutedEventArgs e){if(!loading)InvalidatePreview();}
    private void SelectionChanged(object sender,SelectionChangedEventArgs e){if(!loading)InvalidatePreview();}
    private void IgnoreWheel(object sender,MouseWheelEventArgs e)=>e.Handled=true;
    private void SelectWorkbook(string path)
    {
        workbookPath=path.Trim();
        WorkbookIcon.Source=WindowsFileIcon.Load(".xlsx");
        WorkbookName.Text=string.IsNullOrWhiteSpace(workbookPath)?"No workbook selected":Path.GetFileName(workbookPath);
        WorkbookPath.Text=string.IsNullOrWhiteSpace(workbookPath)?"Excel workbook (.xlsx)":Path.GetDirectoryName(workbookPath);
        WorkbookName.ToolTip=WorkbookPath.ToolTip=workbookPath;
        OpenWorkbookButton.IsEnabled=!IsBusy&&!string.IsNullOrWhiteSpace(workbookPath);
        if(!loading)InvalidatePreview();
    }
    private void Browse_Click(object s,RoutedEventArgs e){var picker=new OpenFileDialog{Title="Choose operations workbook",Filter="Excel workbook (*.xlsx)|*.xlsx"};if(picker.ShowDialog(OwnerWindow)==true)SelectWorkbook(picker.FileName);}
    internal static ProcessStartInfo WorkbookLaunch(string path)
    {
        return WorkbookFile.Launch(path,false);
    }
    private void OpenWorkbook_Click(object sender,RoutedEventArgs e)
    {
        if(IsBusy)return;
        try{Process.Start(WorkbookLaunch(workbookPath));}
        catch(Exception ex){Error(ex);}
    }
    private void Save_Click(object s,RoutedEventArgs e){try{settings=Read();settings.Save();InvalidatePreview();StatusText.Text="Reminder settings saved. Credentials stay in Windows Credential Manager.";}catch(Exception ex){Error(ex);}}
    private void Import_Click(object s,RoutedEventArgs e)
    {
        var picker=new OpenFileDialog{Title="Choose the previous reminder config.json",Filter="JSON settings (*.json)|*.json"};if(picker.ShowDialog(OwnerWindow)!=true)return;
        try{settings=OperationsSettings.Import(picker.FileName);Populate();InvalidatePreview();StatusText.Text="Previous workbook and email settings loaded. Bulk items are enabled. Review and save; passwords were not imported.";}catch(Exception ex){Error(ex);}
    }
    internal void ReloadEmailSettings()
    {
        var saved=OperationsSettings.Load();
        if(settings.SmtpServer==saved.SmtpServer&&settings.SmtpPort==saved.SmtpPort&&settings.From==saved.From&&settings.Subject==saved.Subject&&settings.Recipients.SequenceEqual(saved.Recipients)&&settings.Cc.SequenceEqual(saved.Cc))return;
        settings.SmtpServer=saved.SmtpServer;settings.SmtpPort=saved.SmtpPort;settings.From=saved.From;settings.Subject=saved.Subject;settings.Recipients=saved.Recipients;settings.Cc=saved.Cc;InvalidatePreview();
        StatusText.Text="Email settings changed. Refresh and review the reminder before sending.";
    }
    internal void RefreshCredentials()
    {
        string? building=WindowsCredentialVault.Username(BuildingLinkCredentialStore.Target),email=WindowsCredentialVault.Username(OperationsRunner.EmailTarget);
        CredentialText.Text="BuildingLink: "+(building??"not saved")+" | Reminder email: "+(email??"not saved")+". Saved passwords remain secured for this Windows account.";
    }
    private void Credentials_Click(object s,RoutedEventArgs e){if(IsBusy)return;manageCredentials();RefreshCredentials();}
    private void Busy(bool busy){processing=busy;SetupPanel.IsEnabled=WorkbookPanel.IsEnabled=SchedulePanel.IsEnabled=RefreshButton.IsEnabled=!busy;OpenWorkbookButton.IsEnabled=!busy&&!string.IsNullOrWhiteSpace(workbookPath);SendButton.IsEnabled=EmailPreviewButton.IsEnabled=!busy&&prepared.HasValue;CancelButton.IsEnabled=busy&&!scheduleBusy;if(!busy)Progress.IsIndeterminate=false;}
    private void Receive(JsonElement item)
    {
        Dispatcher.Invoke(()=>{if(item.TryGetProperty("message",out var m))StatusText.Text=m.GetString();if(item.TryGetProperty("percent",out var p)&&p.ValueKind==JsonValueKind.Number){Progress.IsIndeterminate=false;Progress.Value=p.GetDouble();}else Progress.IsIndeterminate=true;});
    }
    private async void Refresh_Click(object s,RoutedEventArgs e)
    {
        if(IsBusy)return;
        try{settings=Read();settings.Save();InvalidatePreview();Busy(true);Progress.Value=0;prepared=await runner.RunAsync("preview",settings,null,Receive);ShowPreview(prepared.Value);SetupExpander.IsExpanded=false;StatusText.Text="Preview ready. No email has been sent.";}
        catch(Exception ex){Error(ex);}finally{Busy(false);RefreshHistory();}
    }
    private void ShowPreview(JsonElement data)
    {
        SummaryText.Text=$"{data.GetProperty("start").GetString()} through {data.GetProperty("end").GetString()} | {data.GetProperty("bookings").GetArrayLength()} bookings | {data.GetProperty("repairs").GetArrayLength()} open repair cases | {data.GetProperty("bulk").GetArrayLength()} open bulk requests";
        WarningText.Text=string.Join("\n",data.GetProperty("warnings").EnumerateArray().Select(v=>v.GetString()));
        if(data.GetProperty("bookings").GetArrayLength()==0&&data.GetProperty("booking_error").ValueKind==JsonValueKind.Null)SummaryText.Text+=". No matching bookings found.";
    }
    private void EmailPreview_Click(object s,RoutedEventArgs e){if(prepared is { } p)ShowText("Reminder email preview","From: "+settings.From+"\nTo: "+string.Join("; ",settings.Recipients)+"\nCc: "+string.Join("; ",settings.Cc)+"\nSubject: "+settings.Subject+"\n\n"+p.GetProperty("body").GetString());}
    private async void Send_Click(object s,RoutedEventArgs e)
    {
        if(IsBusy||prepared is null)return;
        try
        {
            var current=Read();current.Validate(true);
            if(JsonSerializer.Serialize(current)!=JsonSerializer.Serialize(settings))throw new InvalidOperationException("Settings changed. Refresh the preview before sending.");
            if(prepared.Value.GetProperty("warnings").GetArrayLength()>0&&!settings.AllowPartial)throw new InvalidOperationException("Resolve the warnings or explicitly allow reminders with warnings, then refresh the preview.");
            if(MessageBox.Show(OwnerWindow,"Send this reviewed reminder now?\n\nTo: "+string.Join("; ",settings.Recipients)+"\nCc: "+string.Join("; ",settings.Cc)+"\nSubject: "+settings.Subject+"\n\n"+(WarningText.Text.Length>0?"Warnings included:\n"+WarningText.Text:"No preview warnings."),"Send operations reminder",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return;
            var sendPreview=prepared.Value;prepared=null;Busy(true);var result=await runner.RunAsync("send",settings,sendPreview,Receive);StatusText.Text=result.GetProperty("message").GetString();
        }
        catch(Exception ex){Error(ex);}finally{Busy(false);RefreshHistory();}
    }
    private void Cancel_Click(object s,RoutedEventArgs e){Cancel();StatusText.Text="Cancellation requested. A submitted email cannot be recalled; check history before retrying.";}
    private void RefreshHistory()=>HistoryGrid.ItemsSource=new OperationsHistory().Read();
    private void HistoryNavigation_Click(object sender,RoutedEventArgs e)
    {
        if(IsBusy)return;
        bool show=HistoryPanel.Visibility!=Visibility.Visible;
        if(show)RefreshHistory();
        HistoryPanel.Visibility=show?Visibility.Visible:Visibility.Collapsed;
        FeatureScroll.Visibility=ActionCard.Visibility=show?Visibility.Collapsed:Visibility.Visible;
        HistoryNavigationButton.Content=show?"Back to feature":"View history";
    }
    private void History_Click(object s,RoutedEventArgs e){try{RefreshHistory();}catch(Exception ex){Error(ex);}}
    private void RunDetails_Click(object s,RoutedEventArgs e){if(HistoryGrid.SelectedItem is OperationsRun r)ShowText("Reminder run - "+r.LocalTime,r.Action+"\n"+r.Status+"\n\n"+r.Details);}
    private void ShowText(string title,string text)=>new Window{Owner=OwnerWindow,Title=title,Width=850,Height=650,WindowStartupLocation=WindowStartupLocation.CenterOwner,Content=new TextBox{Text=text,IsReadOnly=true,TextWrapping=TextWrapping.Wrap,VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Margin=new Thickness(20)}}.ShowDialog();
    private void Error(Exception ex){StatusText.Text=ex.Message;MessageBox.Show(OwnerWindow,ex.Message,"Reservations & Reminders",MessageBoxButton.OK,MessageBoxImage.Warning);}
    private async Task CheckScheduleAsync(){var r=await OperationsSchedule.ExecuteAsync("Get",settings);ScheduleText.Text=r.GetProperty("Exists").GetBoolean()?(r.GetProperty("Enabled").GetBoolean()?"Enabled. Next run: "+r.GetProperty("NextRun").GetString():"Automatic reminders are disabled."):"No app reminder schedule exists.";var other=r.GetProperty("OtherTasks").EnumerateArray().Select(v=>v.GetString()).ToList();if(other.Count>0)ScheduleText.Text+=" Older/other reminder schedules: "+string.Join(", ",other)+". Disable them in Windows Task Scheduler before enabling this app's schedule.";}
    private async void CheckSchedule_Click(object s,RoutedEventArgs e){if(IsBusy)return;try{scheduleBusy=true;Busy(true);await CheckScheduleAsync();}catch(Exception ex){Error(ex);}finally{scheduleBusy=false;Busy(false);}}
    private async void EnableSchedule_Click(object s,RoutedEventArgs e)
    {
        if(IsBusy)return;
        try
        {
            settings=Read();settings.Validate(true);if(!TimeOnly.TryParseExact(settings.Time,"HH:mm",out _)||settings.Days.Count==0||settings.Days.Any(d=>d<0||d>6))throw new InvalidOperationException("Use HH:mm and choose at least one schedule day.");
            if(WindowsCredentialVault.Username(BuildingLinkCredentialStore.Target) is null||WindowsCredentialVault.Username(OperationsRunner.EmailTarget) is null)throw new InvalidOperationException("Save the required BuildingLink and reminder email credentials first.");
            scheduleBusy=true;Busy(true);var state=await OperationsSchedule.ExecuteAsync("Get",settings);if(state.GetProperty("OtherTasks").GetArrayLength()>0)throw new InvalidOperationException("An older/other reminder schedule is enabled. Disable it in Windows Task Scheduler to prevent duplicate emails.");
            if(MessageBox.Show(OwnerWindow,"Enable automatic reminder emails at "+settings.Time+" using this computer's Windows clock?\n\nTo: "+string.Join("; ",settings.Recipients)+"\nWarnings: "+(settings.AllowPartial?"Reminders with warnings may be sent.":"Sending stops if the preview has warnings.")+"\n\nThis account must be signed in. Settings apply to both manual and scheduled reminders.","Enable automatic reminders",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return;
            settings.Save();InvalidatePreview();await OperationsSchedule.ExecuteAsync("Save",settings);await CheckScheduleAsync();
        }
        catch(Exception ex){Error(ex);}finally{scheduleBusy=false;Busy(false);}
    }
    private async void DisableSchedule_Click(object s,RoutedEventArgs e){if(IsBusy)return;try{scheduleBusy=true;Busy(true);await OperationsSchedule.ExecuteAsync("Disable",settings);await CheckScheduleAsync();}catch(Exception ex){Error(ex);}finally{scheduleBusy=false;Busy(false);}}
    private void TaskScheduler_Click(object s,RoutedEventArgs e)=>Process.Start(new ProcessStartInfo("mmc.exe","taskschd.msc"){UseShellExecute=true});
}
