using System.IO;
using System.Windows;
using Microsoft.Win32;
using PropertyManager.App.BuildingLink;

namespace PropertyManager.App.EmailLists;
public partial class GeneralEmailListView
{
    private RecipientFileData? _mailingSource;
    private string _mailingSourcePath="";
    private bool _previewMailing;
    private bool IncludeMailing=>_ownerMailingMode;
    private void InitializeOwnerMailingMode()
    {
        DefaultCorporationCombo.ItemsSource=OwnerMailingDefault.All;
        if(!_ownerMailingMode)return;
        _updating=true;
        PageTitle.Text="Owner Mailing List";
        IntroDescription.Text="Combine owner records and Unit Data, review mailing addresses, and export your list.";
        GroupHeading.Text="Owner list options";
        ListTypePanel.Visibility=Visibility.Collapsed;ModeCombo.SelectedIndex=(int)EmailListMode.AllOwners;
        IncludeMailingCheck.Visibility=Visibility.Collapsed;IncludeMailingCheck.IsChecked=true;GroupCheck.IsChecked=true;
        DefaultCorporationPanel.Visibility=Visibility.Collapsed;
        AudienceCard.Visibility=Visibility.Collapsed;AudienceWidthColumn.Width=new GridLength(0);AudienceGapColumn.Width=new GridLength(0);
        GenerateOwnerButton.Visibility=OwnerUnitSourceText.Visibility=Visibility.Visible;
        EmailStatusColumn.Visibility=Visibility.Collapsed;
        PreviewGrid.Columns.First(c=>c.Header?.ToString()=="UNIT").DisplayIndex=0;
        NamesColumn.DisplayIndex=1;MailingAddressColumn.DisplayIndex=2;MailingStatusColumn.DisplayIndex=3;
        PreviewGrid.Columns.First(c=>c.Header?.ToString()=="EMAIL (EDITABLE)").DisplayIndex=4;
        MailingNameColumn.DisplayIndex=5;
        var style=new System.Windows.Style(typeof(System.Windows.Controls.DataGridRow));
        var review=new System.Windows.DataTrigger{Binding=new System.Windows.Data.Binding("NeedsMailingReview"),Value=true};
        review.Setters.Add(new System.Windows.Setter(System.Windows.Controls.Control.BackgroundProperty,new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(255,241,240))));
        style.Triggers.Add(review);PreviewGrid.RowStyle=style;
        DefaultAddressHint.Text="DEMO 1001: [unit] - 100 Example Ave, Example City, ON A1A 1A1. DEMO 2002: TH [unit] - 200 Example Ave, Example City, ON A1A 1A2.";
        MailingAddressColumn.Header="MAILING ADDRESS";MailingStatusColumn.Header="ADDRESS SOURCE / REVIEW";
        CopyButton.Visibility=Visibility.Collapsed;ExportButton.Content="Export mailing list";
        PreviewHint.Text="Alternate addresses take priority. Confirmed blanks use the selected corporation's property address. Source backups are never edited.";
        _updating=false;
    }
    private static List<string> CorporationIds(string value)=>System.Text.RegularExpressions.Regex.Matches(value,@"\bDEMO[\s_-]*(1001|2002)\b",System.Text.RegularExpressions.RegexOptions.IgnoreCase).Select(m=>m.Groups[1].Value).Distinct().ToList();
    private void UpdateDefaultCorporation()
    {
        if(!_ownerMailingMode||_source is null)return;
        string folderName=_shortcuts.FirstOrDefault(f=>SameFolder(f.Path,FolderText.Text))?.Name??"";
        var ids=CorporationIds(folderName+" "+string.Join(" ",_source.Rows.Select(r=>GeneratedEmailListService.Value(r,"Building Name"))));
        DefaultCorporationCombo.SelectedItem=ids.Count==1?OwnerMailingDefault.All.First(p=>p.Corporation=="DEMO "+ids[0]):null;
    }
    private void ResetMailingSource()
    {
        _mailingSource=null;_mailingSourcePath="";_previewMailing=false;IncludeMailingCheck.IsChecked=_ownerMailingMode;DefaultCorporationCombo.SelectedItem=null;
        MailingNameColumn.Visibility=MailingAddressColumn.Visibility=MailingStatusColumn.Visibility=Visibility.Collapsed;
        UpdateMailingControls();
    }
    private void UpdateMailingControls()
    {
        OwnerMailingPanel.Visibility=_ownerMailingMode?Visibility.Visible:Visibility.Collapsed;
        MailingSourcePanel.Visibility=IncludeMailing?Visibility.Visible:Visibility.Collapsed;
        MailingSourceText.Text=_mailingSource is not null?"Unit Data: "+Path.GetFileName(_mailingSourcePath):_source is not null&&AlternateMailingAddress.HasColumns(_source)
            ?"Alternate-address columns found in occupant data."
            :"Alternate-address columns are not in this occupant backup. Choose its matching Unit Data backup.";
        if(_mailingSource is not null&&BackupDateFromFileName(_mailingSourcePath) is {} unitDate&&BackupDateFromFileName(_loadedFile) is {} occupantDate&&unitDate.Date!=occupantDate.Date)
            MailingSourceText.Text+=" Different backup dates: review this source pairing.";
        OwnerUnitSourceText.Text=_mailingSource is not null?MailingSourceText.Text:"Unit Data: select a corporation with both backup folders saved in Manage folders. A matching backup is needed before generating.";
    }
    private async void MailingOption_Changed(object s,RoutedEventArgs e)
    {
        if(_updating||!IsInitialized)return;
        UpdateMailingControls();MarkFiltersChanged();
        if(IncludeMailing&&_source is not null&&_mailingSource is null&&!AlternateMailingAddress.HasColumns(_source))await TryLoadMatchingUnitDataAsync();
    }
    internal static string? FindMatchingUnitBackup(string folder,DateTime occupantDate)
    {
        return Directory.EnumerateFiles(folder).Where(p=>!Path.GetFileName(p).StartsWith("~$")&&Path.GetFileName(p).Contains("Unit-Data",StringComparison.OrdinalIgnoreCase)&&new[]{".xls",".xlsx",".xlsm",".csv"}.Contains(Path.GetExtension(p),StringComparer.OrdinalIgnoreCase))
            .Select(p=>new{Path=p,Date=BackupDateFromFileName(p)??File.GetLastWriteTime(p)})
            .Where(f=>f.Date.Date==occupantDate.Date&&Math.Abs((f.Date-occupantDate).TotalHours)<=24)
            .OrderBy(f=>Math.Abs((f.Date-occupantDate).TotalSeconds)).ThenByDescending(f=>f.Date).ThenBy(f=>f.Path,StringComparer.OrdinalIgnoreCase).Select(f=>f.Path).FirstOrDefault();
    }
    private async Task TryLoadMatchingUnitDataAsync()
    {
        try
        {
            string folder=_shortcuts.FirstOrDefault(f=>SameFolder(f.Path,FolderText.Text))?.UnitPath??"";
            if(string.IsNullOrWhiteSpace(folder))
            {
                var destinations=BuildingLinkSettings.Load().Buildings.Where(b=>SameFolder(b.OccupantFolder,FolderText.Text)).ToList();
                if(destinations.Count==1)folder=destinations[0].UnitFolder;
            }
            if(string.IsNullOrWhiteSpace(folder)){OwnerUnitSourceText.Text="Unit Data folder is missing. Add it under Manage folders, then Refresh.";return;}
            SetBusy(true);
            var date=BackupDateFromFileName(_loadedFile)??File.GetLastWriteTime(_loadedFile);
            string? path=await Task.Run(()=>FindMatchingUnitBackup(folder,date));
            if(path is not null)await LoadUnitBackupAsync(path);
            else OwnerUnitSourceText.Text="No matching Unit Data backup for this occupant backup date. Refresh after downloading both reports, or choose Unit Data in Backup & source details.";
        }
        catch(OperationCanceledException){}
        catch(Exception ex)when(IsFileError(ex)){MailingSourceText.Text="Matching Unit Data could not be loaded. Choose a backup manually. "+ex.Message;OwnerUnitSourceText.Text=MailingSourceText.Text;}
        finally{SetBusy(false);}
    }
    private async void ChooseUnitData_Click(object s,RoutedEventArgs e)
    {
        if(_busy||_source is null)return;
        var picker=new OpenFileDialog{Title="Choose matching Unit Data for this corporation",Filter="Unit Data backups (*.xls;*.xlsx;*.xlsm;*.csv)|*.xls;*.xlsx;*.xlsm;*.csv",CheckFileExists=true,Multiselect=false};
        if(_mailingSourcePath.Length>0)picker.InitialDirectory=Path.GetDirectoryName(_mailingSourcePath);
        if(picker.ShowDialog(OwnerWindow)!=true)return;
        try{SetBusy(true);await LoadUnitBackupAsync(picker.FileName);}
        catch(OperationCanceledException){}
        catch(Exception ex)when(IsFileError(ex)){ShowError(ex);}
        finally{SetBusy(false);}
    }
    private async Task LoadUnitBackupAsync(string path)
    {
        var loaded=await RecipientFileLoader.LoadAsync(path,OwnerWindow,"Unit Data");
        await Task.Run(()=>AlternateMailingAddress.UnitMap(loaded.Data));
        var occupantBuildings=_source!.Rows.Select(r=>GeneratedEmailListService.Value(r,"Building Name")).Where(v=>v.Length>0).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if(loaded.Data.Columns.Contains("Building Name",StringComparer.OrdinalIgnoreCase)&&loaded.Data.Rows.Any(r=>GeneratedEmailListService.Value(r,"Building Name") is string name&&name.Length>0&&!occupantBuildings.Contains(name)))
            throw new InvalidDataException("The Unit Data building does not match the occupant backup. Choose the same corporation's export.");
        _mailingSource=loaded.Data;_mailingSourcePath=path;UpdateMailingControls();MarkFiltersChanged();
    }
}
