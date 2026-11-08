using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using PropertyManager.App.EmailLists;
using PropertyManager.App.Services;
using Microsoft.Win32;

namespace PropertyManager.App.BuildingLink;

public partial class BuildingLinkView : UserControl
{
    private BuildingLinkSettings settings=new();
    private readonly BuildingLinkRunner runner=new();
    private readonly Action manageCredentials;
    private readonly Stopwatch elapsed=new();
    private readonly DispatcherTimer timer=new(){Interval=TimeSpan.FromSeconds(1)};
    private readonly List<CheckBox> days=[];
    private List<string> otherTasks=[];
    private bool scheduleBusy;
    private DateTime lastResultWrite;
    private int completedReports, totalReports;
    internal bool IsRunning=>runner.IsRunning;
    internal bool ScheduleBusy=>scheduleBusy;
    public void Cancel()=>runner.Cancel();
    internal BuildingLinkView(Action manageCredentials)
    {
        this.manageCredentials=manageCredentials;
        InitializeComponent();
        try {
            settings=BuildingLinkSettings.Load();
        }catch(Exception ex){FolderStatus.Text="Settings could not be loaded. "+ex.Message;DownloadButton.IsEnabled=false;}
        CorporationItems.ItemsSource=settings.Buildings;
        ScheduleTime.Text=settings.Time;
        foreach(int d in new[]{1,2,3,4,5,6,0})
        {
            var day=new CheckBox{Content=((DayOfWeek)d).ToString()[..3],Tag=d,IsChecked=settings.Days.Contains(d),Margin=new Thickness(0,0,18,8)};
            days.Add(day);DaysPanel.Children.Add(day);
        }
        timer.Tick+=(_,_)=> {if(runner.IsRunning) UpdateProgressText();else if(IsVisible) RefreshResult();};
        Loaded+=async(_,_)=> {timer.Start();RefreshCredentials();RefreshResult();await RefreshScheduleAsync();};
        Unloaded+=(_,_)=>timer.Stop();
    }
    private Window OwnerWindow=>Window.GetWindow(this);
    private void Error(Exception ex)=>MessageBox.Show(OwnerWindow,ex.Message,"BuildingLink Data Download",MessageBoxButton.OK,MessageBoxImage.Warning);
    internal void RefreshCredentials()
    {
        try {var username=WindowsCredentialVault.Username(BuildingLinkCredentialStore.Target);CredentialText.Text=username is null?"No saved BuildingLink login. Add one in Manage credentials.":"Using saved login: "+username+" · Password saved";}
        catch(Exception){CredentialText.Text="Windows could not read the saved BuildingLink login.";}
    }
    private void Credentials_Click(object sender,RoutedEventArgs e){manageCredentials();RefreshCredentials();}
    private void Browse(object sender,bool occupant)
    {
        if(((FrameworkElement)sender).DataContext is not BuildingDestination b)return;
        var picker=new OpenFolderDialog{Title="Choose "+b.Corporation+" "+(occupant?"Occupant Data":"Unit Data")+" folder",Multiselect=false};
        if(picker.ShowDialog(OwnerWindow)!=true)return;
        if(occupant)b.OccupantFolder=picker.FolderName;else b.UnitFolder=picker.FolderName;
    }
    private void BrowseOccupant_Click(object s,RoutedEventArgs e)=>Browse(s,true);
    private void BrowseUnit_Click(object s,RoutedEventArgs e)=>Browse(s,false);
    private void Open(object sender,bool occupant)
    {
        if(((FrameworkElement)sender).DataContext is not BuildingDestination b)return;
        try {string path=occupant?b.OccupantFolder:b.UnitFolder;if(!Directory.Exists(path))throw new DirectoryNotFoundException("This folder does not exist or is unavailable.");var info=new ProcessStartInfo("explorer.exe"){UseShellExecute=true};info.ArgumentList.Add(path);Process.Start(info);}catch(Exception ex){Error(ex);}
    }
    private void OpenOccupant_Click(object s,RoutedEventArgs e)=>Open(s,true);
    private void OpenUnit_Click(object s,RoutedEventArgs e)=>Open(s,false);
    private void SaveFolders_Click(object s,RoutedEventArgs e){try{SaveDownloadSettings(false);FolderStatus.Text="Download selection saved. Folder paths are shared by all features.";}catch(Exception ex){Error(ex);}}
    private void SaveDownloadSettings(bool requireDownloads=true)
    {
        var latest=BuildingLinkSettings.Load();
        foreach(var b in latest.Buildings){var choice=settings.Buildings.FirstOrDefault(v=>v.Corporation==b.Corporation);if(choice is not null)b.Enabled=choice.Enabled;}
        latest.Time=settings.Time;latest.Days=settings.Days;if(requireDownloads)latest.Validate();latest.Save();settings=latest;CorporationItems.ItemsSource=settings.Buildings;
    }
    internal void ReloadSharedFolders()
    {
        if(runner.IsRunning||scheduleBusy)return;
        try{settings=BuildingLinkSettings.Load();CorporationItems.ItemsSource=settings.Buildings;DownloadButton.IsEnabled=true;FolderStatus.Text="Shared folders are used by all occupant-data features.";}
        catch(Exception ex){FolderStatus.Text=ex.Message;DownloadButton.IsEnabled=false;}
    }
    private void ManageFolders_Click(object s,RoutedEventArgs e)
    {
        if(CorporationConfiguration.Manage(OwnerWindow)){ReloadSharedFolders();FolderStatus.Text="Corporation folders saved for all features.";}
    }
    private async void Download_Click(object s,RoutedEventArgs e)
    {
        if(runner.IsRunning || scheduleBusy)return;
        bool invoked=false;
        try {
            settings.Validate();SaveDownloadSettings();
            if(WindowsCredentialVault.Username(BuildingLinkCredentialStore.Target) is null)throw new InvalidOperationException("Save a BuildingLink login in Manage credentials first.");
            SetupPanel.IsEnabled=DownloadButton.IsEnabled=false;CancelButton.IsEnabled=true;elapsed.Restart();
            completedReports=0;totalReports=settings.Buildings.Count(b=>b.Enabled)*2;
            DownloadProgress.Value=0;DownloadProgress.IsIndeterminate=true;StageText.Text="Starting hidden BuildingLink download...";
            UpdateProgressText();
            ResultText.Clear();
            invoked=true;
            int exit=await runner.RunAsync(data=>ApplyProgress(data));
            if(exit!=0 && !File.Exists(BuildingLinkSettings.ResultPath))StageText.Text="Download stopped before completion. Check download components and settings.";
        }catch(Exception ex){StageText.Text=ex.Message;}
        finally{elapsed.Stop();SetupPanel.IsEnabled=DownloadButton.IsEnabled=true;CancelButton.IsEnabled=false;DownloadProgress.IsIndeterminate=false;if(invoked){lastResultWrite=DateTime.MinValue;RefreshResult();}UpdateProgressText();}
    }
    private void Cancel_Click(object s,RoutedEventArgs e){runner.Cancel();CancelButton.IsEnabled=false;StageText.Text="Stopping at the next safe step...";}
    private void UpdateProgressText(bool background=false)
    {
        string time=background?"":$" - {elapsed.Elapsed:hh\\:mm\\:ss} elapsed";
        ProgressText.Text=$"{DownloadProgress.Value:0}%{time} - {completedReports} of {totalReports} reports saved"+(background?" - Background download":runner.IsRunning?" - Cancel stops at the next safe step.":"");
    }
    internal void ApplyProgress(JsonElement data,bool background=false)
    {
        if(data.TryGetProperty("message",out var message))StageText.Text=message.GetString();
        if(data.TryGetProperty("completed",out var done))completedReports=done.GetInt32();
        else if(data.TryGetProperty("files",out var files))completedReports=files.GetArrayLength();
        if(data.TryGetProperty("total",out var total))totalReports=total.GetInt32();
        bool hasMilestone=data.TryGetProperty("progress_percent",out var milestone)&&milestone.TryGetDouble(out double value)&&double.IsFinite(value);
        DownloadProgress.Value=hasMilestone?Math.Clamp(milestone.GetDouble(),0,100):totalReports>0?Math.Clamp(completedReports*100.0/totalReports,0,100):0;
        bool finished=data.TryGetProperty("type",out var type)&&type.GetString()=="complete" || data.TryGetProperty("finished",out _);
        DownloadProgress.IsIndeterminate=!finished&&!hasMilestone&&completedReports==0;
        UpdateProgressText(background);
    }
    internal void RefreshResult()
    {
        if(File.Exists(BuildingLinkSettings.SuccessPath))
        {
            try {using var success=JsonDocument.Parse(File.ReadAllText(BuildingLinkSettings.SuccessPath));
                if(DateTimeOffset.TryParse(success.RootElement.GetProperty("finished").GetString(),out var date))LastSuccessText.Text="Last successful download: "+date.ToLocalTime().ToString("g");}
            catch(Exception ex) when(ex is IOException or JsonException or KeyNotFoundException or InvalidOperationException) { }
        }
        if(!runner.IsRunning && File.Exists(BuildingLinkSettings.ProgressPath))
        {
            try
            {
                using var live=JsonDocument.Parse(File.ReadAllText(BuildingLinkSettings.ProgressPath));var state=live.RootElement;
                if(state.GetProperty("running").GetBoolean() && DateTimeOffset.TryParse(state.GetProperty("updated").GetString(),out var updated) && DateTimeOffset.UtcNow-updated<TimeSpan.FromMinutes(5))
                {
                    ApplyProgress(state,background:true);
                    DownloadButton.IsEnabled=false;return;
                }
            }
            catch(Exception ex) when(ex is IOException or JsonException or KeyNotFoundException or InvalidOperationException) { }
        }
        if(!runner.IsRunning){DownloadProgress.IsIndeterminate=false;if(!scheduleBusy)DownloadButton.IsEnabled=true;}
        if(!File.Exists(BuildingLinkSettings.ResultPath))return;
        try {
            DateTime write=File.GetLastWriteTimeUtc(BuildingLinkSettings.ResultPath);if(write==lastResultWrite)return;
            using var data=JsonDocument.Parse(File.ReadAllText(BuildingLinkSettings.ResultPath));var result=data.RootElement;
            string message=result.GetProperty("message").GetString()??"";
            StageText.Text=message;
            LastRunText.Text="Last download: "+(result.TryGetProperty("finished",out var finished)&&DateTimeOffset.TryParse(finished.GetString(),out var date)?date.ToLocalTime().ToString("g"):"unknown time")+" · "+(result.GetProperty("success").GetBoolean()?"Successful":"Needs attention");
            var lines=new List<string>();int count=0;
            if(result.TryGetProperty("files",out var files))foreach(var file in files.EnumerateArray()){lines.Add(file.GetProperty("path").GetString()??"");count++;}
            if(result.TryGetProperty("failures",out var failures))foreach(var failure in failures.EnumerateArray())lines.Add(failure.GetProperty("corporation").GetString()+": "+failure.GetProperty("reason").GetString());
            ResultText.Text=string.Join(Environment.NewLine+Environment.NewLine,lines);
            totalReports=settings.Buildings.Count(b=>b.Enabled)*2;completedReports=count;ApplyProgress(result);lastResultWrite=write;
        }catch(IOException){}catch(JsonException){}catch(InvalidOperationException){}catch(KeyNotFoundException){}
    }
    private async Task RefreshScheduleAsync()
    {
        try {
            var data=await BuildingLinkSchedule.ExecuteAsync("Get",settings);
            otherTasks=data.GetProperty("OtherTasks").EnumerateArray().Select(v=>v.GetString()??"").ToList();
            bool enabled=data.GetProperty("Enabled").GetBoolean();
            ScheduleStatus.Text=enabled?"Enabled · next run: "+data.GetProperty("NextRun").GetString():"Automatic updates are disabled.";
            if(otherTasks.Count>0)ScheduleStatus.Text+=" Other downloader schedules found: "+string.Join(", ",otherTasks)+". Review them before enabling another schedule.";
            DisableScheduleButton.IsEnabled=data.GetProperty("Exists").GetBoolean()&&enabled;
        }catch(Exception ex){ScheduleStatus.Text="Schedule status unavailable. "+ex.Message[..Math.Min(350,ex.Message.Length)];}
    }
    private async void RefreshSchedule_Click(object s,RoutedEventArgs e)=>await RefreshScheduleAsync();
    private async void EnableSchedule_Click(object s,RoutedEventArgs e)
    {
        if(scheduleBusy||runner.IsRunning)return;
        try {
            scheduleBusy=true;SetupPanel.IsEnabled=DownloadButton.IsEnabled=false;
            settings.Validate();
            if(!TimeOnly.TryParseExact(ScheduleTime.Text.Trim(),"HH:mm",System.Globalization.CultureInfo.InvariantCulture,System.Globalization.DateTimeStyles.None,out _))throw new InvalidOperationException("Enter a time in HH:mm format, such as 06:00.");
            settings.Time=ScheduleTime.Text.Trim();settings.Days=days.Where(d=>d.IsChecked==true).Select(d=>(int)d.Tag).ToList();
            if(settings.Days.Count==0)throw new InvalidOperationException("Select at least one schedule day.");
            if(WindowsCredentialVault.Username(BuildingLinkCredentialStore.Target) is null)throw new InvalidOperationException("Save a BuildingLink login before enabling automatic downloads.");
            await RefreshScheduleAsync();
            if(otherTasks.Count>0&&MessageBox.Show(OwnerWindow,"Other download schedules are enabled. Creating another can cause duplicate downloads. Continue without changing those tasks?","Existing downloader schedules",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return;
            scheduleBusy=true;SetupPanel.IsEnabled=DownloadButton.IsEnabled=false;SaveDownloadSettings();
            await BuildingLinkSchedule.ExecuteAsync("Save",settings);await RefreshScheduleAsync();
        }catch(Exception ex){Error(ex);}finally{scheduleBusy=false;SetupPanel.IsEnabled=DownloadButton.IsEnabled=true;}
    }
    private async void DisableSchedule_Click(object s,RoutedEventArgs e)
    {
        if(scheduleBusy)return;
        try {scheduleBusy=true;SetupPanel.IsEnabled=DownloadButton.IsEnabled=false;await BuildingLinkSchedule.ExecuteAsync("Disable",settings);await RefreshScheduleAsync();}
        catch(Exception ex){Error(ex);}finally{scheduleBusy=false;SetupPanel.IsEnabled=DownloadButton.IsEnabled=true;}
    }
}
