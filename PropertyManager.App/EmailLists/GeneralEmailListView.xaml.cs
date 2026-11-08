using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;

namespace PropertyManager.App.EmailLists;

public partial class GeneralEmailListView : UserControl
{
    private readonly bool _ownerMailingMode;
    private RecipientFileData? _source;
    private List<GeneratedRecipient> _preview = [];
    private List<FolderShortcut> _shortcuts = [];
    private bool _updating, _busy, _previewEdited, _filtersDirty;
    internal bool IsBusy=>_busy;
    private string _loadedFile = "";
    internal sealed record BackupFile(string Path, DateTime Modified)
    {
        public DateTime BackupDate => BackupDateFromFileName(Path) ?? Modified;
        public string DisplayName => $"{BackupDate:MMM d, yyyy h:mm tt} — {System.IO.Path.GetFileName(Path)}";
    }
    public GeneralEmailListView(bool ownerMailingMode=false)
    {
        _ownerMailingMode=ownerMailingMode;
        InitializeComponent();
        InitializeOwnerMailingMode();if(_ownerMailingMode)UseCampaignButton.Visibility=Visibility.Collapsed;
        SizeChanged += (_, _) => UpdateCompactLayout();
        foreach (var column in PreviewGrid.Columns.OfType<DataGridTextColumn>())
        {
            Style style = new(typeof(TextBlock));
            style.Setters.Add(new Setter(TextBlock.TextWrappingProperty, TextWrapping.Wrap));
            style.Setters.Add(new Setter(TextBlock.FontSizeProperty, 12.0));
            if (column.Binding is System.Windows.Data.Binding binding) style.Setters.Add(new Setter(ToolTipProperty, new System.Windows.Data.Binding(binding.Path.Path)));
            column.ElementStyle = style;
        }
        UpdateFilterControls();
        try { _shortcuts = FolderShortcutStore.Load(); }
        catch (Exception ex) when (IsFileError(ex)) { SourceText.Text = "Saved folders could not be loaded. " + ex.Message; }
        RenderShortcuts();
    }
    private void UpdateCompactLayout()
    {
        bool compact = ActualHeight > 0 && ActualHeight < 700;
        IntroDescription.Visibility = SourceMetadata.Visibility = PreviewHint.Visibility = compact ? Visibility.Collapsed : Visibility.Visible;
        BackupDetails.Header = compact ? BackupDateText.Text : "Backup & source details";
    }
    private Window OwnerWindow => Window.GetWindow(this);
    private static bool IsFileError(Exception ex) => ex is IOException or UnauthorizedAccessException or ArgumentException or System.Xml.XmlException or System.Text.Json.JsonException or InvalidOperationException or FormatException or Microsoft.VisualBasic.FileIO.MalformedLineException or System.Text.DecoderFallbackException;
    private void ShowError(Exception ex) => MessageBox.Show(OwnerWindow, ex.Message, PageTitle.Text, MessageBoxButton.OK, MessageBoxImage.Warning);
    public bool ConfirmDiscard()
    {
        if (!_previewEdited) return true;
        if (MessageBox.Show(OwnerWindow, "This will replace your edited recipient preview. Continue?", "Unsaved preview edits", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes) return false;
        _previewEdited = false;
        return true;
    }
    private void SetBusy(bool busy)
    {
        _busy = busy;
        UpdateCompactLayout();
        ShortcutButtons.IsEnabled = ManageFoldersButton.IsEnabled = !busy;
        RefreshFolderButton.IsEnabled = !busy && FolderText.Text.Length > 0;
        ChooseBackupButton.IsEnabled = !busy && FolderText.Text.Length > 0;
        FilterPanel.IsEnabled = !busy && _source is not null;
        LoadProgress.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
    }
    internal static List<BackupFile> FindBackups(string folder) => Directory.EnumerateFiles(folder)
        .Where(p => !System.IO.Path.GetFileName(p).StartsWith("~$") && new[] { ".xlsx", ".xlsm", ".csv" }.Contains(System.IO.Path.GetExtension(p), StringComparer.OrdinalIgnoreCase))
        .Select(p => new BackupFile(p, File.GetLastWriteTime(p))).OrderByDescending(f => f.BackupDate).ThenByDescending(f => f.Modified).ThenBy(f => f.Path, StringComparer.OrdinalIgnoreCase).ToList();
    private async void RefreshFolder_Click(object sender, RoutedEventArgs e) { if (ConfirmDiscard()) await LoadFolderAsync(); }
    private async void ChooseFolder_Click(object sender, RoutedEventArgs e)
    {
        OpenFolderDialog picker = new() { Title = "Choose BuildingLink occupant backup folder", Multiselect = false };
        if (picker.ShowDialog(OwnerWindow) != true || !ConfirmDiscard()) return;
        var existing = _shortcuts.FirstOrDefault(f => SameFolder(f.Path, picker.FolderName));
        if (existing is not null) { await SelectFolderAsync(existing.Path); return; }
        FolderShortcutEditWindow editor = new(new(SuggestFolderName(picker.FolderName), picker.FolderName), _shortcuts) { Owner = OwnerWindow };
        if (editor.ShowDialog() == true && editor.SavedFolder is { } saved && SaveShortcuts([.. _shortcuts, saved])) await SelectFolderAsync(saved.Path);
    }
    private async Task SelectFolderAsync(string path)
    {
        FolderText.Text = path;
        await LoadFolderAsync();
    }
    private async Task LoadFolderAsync()
    {
        if (_busy || FolderText.Text.Length == 0) return;
        ResetSource();
        SetBusy(true);
        RenderShortcuts();
        try
        {
            string folderPath = FolderText.Text;
            var files = await Task.Run(() => FindBackups(folderPath));
            var latest = files.FirstOrDefault(f => System.IO.Path.GetFileName(f.Path).Contains("Occupant", StringComparison.OrdinalIgnoreCase)) ?? files.FirstOrDefault();
            if (latest is null) { BackupDateText.Text = "No supported backups found."; SourceText.Text = "Choose a folder containing BuildingLink .xlsx, .xlsm or .csv exports."; }
            else await LoadBackupAsync(latest);
        }
        catch (Exception ex) when (IsFileError(ex)) { ResetSource(); SourceText.Text = "The backup folder is unavailable. Check the path and network access."; ShowError(ex); }
        finally { _updating = false; SetBusy(false); }
    }
    internal OpenFileDialog CreateBackupPicker() => new()
    {
        Title = "Choose an older occupant backup",
        InitialDirectory = FolderText.Text,
        Filter = "Occupant backups (*.xlsx;*.xlsm;*.csv)|*.xlsx;*.xlsm;*.csv|Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|CSV files (*.csv)|*.csv",
        CheckFileExists = true,
        Multiselect = false,
        RestoreDirectory = true
    };
    private async void ChooseBackup_Click(object sender, RoutedEventArgs e)
    {
        if (_busy) return;
        try
        {
            var picker = CreateBackupPicker();
            if (picker.ShowDialog(OwnerWindow) == true) await SelectBackupAsync(picker.FileName);
        }
        catch (Exception ex) when (IsFileError(ex)) { ShowError(ex); }
    }
    private async Task SelectBackupAsync(string path)
    {
        if (_busy || !ConfirmDiscard()) return;
        SetBusy(true);
        try { await LoadBackupAsync(new BackupFile(path, File.GetLastWriteTime(path))); }
        catch (Exception ex) when (IsFileError(ex)) { ShowError(ex); }
        finally { SetBusy(false); }
    }
    private async Task LoadBackupAsync(BackupFile file)
    {
        ResetSource();
        SelectedFileText.Text = file.Path;
        BackupDateText.Text = $"Backup: {file.BackupDate:MMM d, yyyy · h:mm tt}";
        try
        {
            var loaded = await RecipientFileLoader.LoadAsync(file.Path, OwnerWindow, "Occupants Data");
            await Task.Run(() => GeneratedEmailListService.Validate(loaded.Data));
            _source = loaded.Data;
            _loadedFile = file.Path;
            WorksheetText.Text = "Worksheet: " + loaded.Sheet;
            SourceText.Text = $"{_source.Rows.Count:N0} occupants loaded · {string.Join(", ", _source.Rows.Select(r => GeneratedEmailListService.Value(r, "Building Name")).Distinct())}";
            UpdateDefaultCorporation();
            UpdateFilterControls();
            if(_ownerMailingMode)await TryLoadMatchingUnitDataAsync();
            Generate();
        }
        catch (OperationCanceledException) { SourceText.Text = "Worksheet selection cancelled. Choose the backup again to retry."; }
        catch (Exception ex) when (IsFileError(ex)) { SourceText.Text = "The selected file could not be loaded as BuildingLink occupant data."; ShowError(ex); }
    }
    private void ResetSource()
    {
        _source = null; _loadedFile = "";
        ResetUnitSelection();ResetMailingSource();
        _preview = []; _previewEdited = _filtersDirty = false;
        PreviewGrid.ItemsSource = null;
        CopyButton.IsEnabled = ExportButton.IsEnabled = false;
        EmptyPreview.Visibility = Visibility.Visible;
        SummaryText.Text = "Choose a corporation and backup to begin.";
        SelectedFileText.Clear(); WorksheetText.Text = ""; ValuesList.ItemsSource = null;
        BackupDateText.Text = "No backup loaded.";
    }
    private EmailListMode Mode => (EmailListMode)ModeCombo.SelectedIndex;
    private bool UseLine => RiserCombo.SelectedIndex == 1;
    private void Dropdown_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
    {
        if (sender is ComboBox { IsDropDownOpen: true }) return;
        e.Handled = true;
        AudienceScroll.ScrollToVerticalOffset(AudienceScroll.VerticalOffset - e.Delta / 120.0 * 48);
    }
    private void Filter_Changed(object sender, RoutedEventArgs e)
    {
        if (_updating || !IsInitialized) return;
        if (!_ownerMailingMode)
        {
            bool grouped = GroupCheck.IsChecked == true;
            _updating = true;
            if (grouped) DeduplicateCheck.IsChecked = false;
            DeduplicateCheck.IsEnabled = !grouped;
            DeduplicateCheck.ToolTip = grouped
                ? "Shared addresses are kept for each unit. Repeated addresses within the same unit are included once."
                : "Remove repeated email addresses across the list.";
            _updating = false;
        }
        if (ReferenceEquals(sender, ModeCombo)) UpdateFilterControls();
        UpdateMailingControls();MarkFiltersChanged();
    }
    private void Riser_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_updating || !IsInitialized) return;
        UpdateFilterControls(); MarkFiltersChanged();
    }
    private void MarkFiltersChanged()
    {
        if (_source is null) return;
        _filtersDirty = true;
        CopyButton.IsEnabled = ExportButton.IsEnabled = false;
        SummaryText.Text = "Filters changed. Generate preview to apply them. Any preview edits are retained until then.";
    }
    private bool UsesBlocks
    {
        get
        {
            string corporation = _shortcuts.FirstOrDefault(f => SameFolder(f.Path, FolderText.Text))?.Name
                ?? (_source is null ? "" : string.Join(" ", _source.Rows.Select(r => GeneratedEmailListService.Value(r, "Building Name"))));
            return CorporationIds(corporation).SequenceEqual(new[] { "2002" });
        }
    }
    private void UpdateFilterControls()
    {
        bool blocks = UsesBlocks;
        ((ComboBoxItem)ModeCombo.Items[1]).Content = blocks ? "Blocks" : "Floors";
        ValuesHeading.Text = Mode == EmailListMode.Floors ? (blocks ? "Select blocks" : "Select floors") : "Select risers";
        FloorColumn.Header = blocks ? "BLOCK" : "FLOOR";
        bool values = Mode is EmailListMode.Risers or EmailListMode.Floors;
        AudiencePanel.Visibility = values || Mode == EmailListMode.Units ? Visibility.Visible : Visibility.Collapsed;
        ValuesPanel.Visibility = values ? Visibility.Visible : Visibility.Collapsed;
        UnitsPanel.Visibility = Mode == EmailListMode.Units ? Visibility.Visible : Visibility.Collapsed;
        RefreshUnitSuggestions();UpdateMailingControls();
        RiserPanel.Visibility = Mode == EmailListMode.Risers ? Visibility.Visible : Visibility.Collapsed;
        _updating = true;
        ValuesList.ItemsSource = _source is null || !values ? null : _source.Rows.Select(r => Mode == EmailListMode.Floors ? GeneratedEmailListService.Value(r, "Floor") : GeneratedEmailListService.Riser(r, UseLine))
            .Where(v => v.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(v => int.TryParse(v, out int n) ? n : int.MaxValue).ThenBy(v => v, StringComparer.OrdinalIgnoreCase).ToList();
        _updating = false;
        RuleText.Text = Mode switch {
            EmailListMode.Units => "Choose individual units, then an audience. Active-user and unit filters still apply.",
            EmailListMode.Residents => "Residential owners, residents, renters, tenants and family members. Off-site owners and retail contacts are excluded.",
            EmailListMode.OnsiteOwners => "Residential contacts labelled Owner: owners living in their own units.",
            EmailListMode.AllOwners => "Owner and Offsite Owner contacts, including retail owners.",
            EmailListMode.Retail => "Contacts whose unit begins with Retail, including off-site owners.",
            EmailListMode.Floors => blocks ? "Select one or more townhouse blocks (100, 200, 300...) and an audience. Uses the block values in BuildingLink's Floor field." : "Uses BuildingLink's Floor field. Select one or more floors and an audience.",
            _ => UseLine ? "Uses BuildingLink Line exactly. Blank lines and retail units have no riser." : "Uses the last two unit digits, including PH units. Confirm this matches the building's riser layout."
        };
    }
    private void Generate_Click(object sender, RoutedEventArgs e) => Generate();
    private void Generate()
    {
        if (_source is null) return;
        if(_ownerMailingMode)
        {
            var property=DefaultCorporationCombo.SelectedItem as OwnerMailingDefault;
            var identities=CorporationIds(string.Join(" ",_source.Rows.Select(r=>GeneratedEmailListService.Value(r,"Building Name"))));
            if(property is not null&&identities.Count==1&&property.Corporation!="DEMO "+identities[0]){MarkFiltersChanged();SummaryText.Text="The selected corporation does not match the occupant backup.";return;}
            if(_mailingSource is null){MarkFiltersChanged();SummaryText.Text="Load matching Unit Data to check alternate addresses before using property defaults.";return;}
        }
        if (Mode is EmailListMode.Floors or EmailListMode.Risers && ValuesList.SelectedItems.Count == 0) { MarkFiltersChanged(); SummaryText.Text = Mode == EmailListMode.Floors ? $"Select at least one {(UsesBlocks ? "block" : "floor")}, then generate the preview." : "Select at least one riser, then generate the preview."; return; }
        if (Mode == EmailListMode.Units && (_selectedUnits.Count == 0 || !string.IsNullOrWhiteSpace(UnitSearch.Text)))
        { MarkFiltersChanged(); SummaryText.Text = "Select a matching unit with Enter or a click, then generate the preview."; return; }
        if (!ConfirmDiscard()) return;
        _preview = GeneratedEmailListService.Generate(_source, Mode, (EmailListAudience)AudienceCombo.SelectedIndex, Mode == EmailListMode.Units ? _selectedUnits : ValuesList.SelectedItems.Cast<string>(), UseLine, _ownerMailingMode||ActiveCheck.IsChecked == true, _ownerMailingMode||DeduplicateCheck.IsChecked == true, _ownerMailingMode||GroupCheck.IsChecked == true, IncludeMailing, _mailingSource, _ownerMailingMode?((DefaultCorporationCombo.SelectedItem as OwnerMailingDefault)?.Corporation??""):null, _ownerMailingMode);
        _previewMailing=IncludeMailing;
        MailingNameColumn.Visibility=MailingAddressColumn.Visibility=MailingStatusColumn.Visibility=_previewMailing?Visibility.Visible:Visibility.Collapsed;
        _previewEdited = _filtersDirty = false;
        NamesColumn.Visibility = _ownerMailingMode||GroupCheck.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        FirstNameColumn.Visibility = LastNameColumn.Visibility = _ownerMailingMode||GroupCheck.IsChecked == true ? Visibility.Collapsed : Visibility.Visible;
        PreviewGrid.ItemsSource = _preview;
        EmptyPreview.Visibility = _preview.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        RefreshSummary();
    }
    private void RefreshSummary()
    {
        if (_filtersDirty) { MarkFiltersChanged(); return; }
        if(_ownerMailingMode||GroupCheck.IsChecked==true){foreach(var row in _preview)GeneratedEmailListService.RefreshEmails([row],true);}else GeneratedEmailListService.RefreshEmails(_preview, DeduplicateCheck.IsChecked == true);
        int unique = _preview.SelectMany(r => GeneratedEmailListService.EmailParts(r.Email)).Where(GeneratedEmailListService.IsValidEmail).Distinct(StringComparer.OrdinalIgnoreCase).Count();
        SummaryText.Text = $"{_preview.Count:N0} {(GroupCheck.IsChecked == true ? (_previewMailing?"unit/address groups":"units") : "contacts")} · {unique:N0} unique valid addresses · {_preview.Count(r => r.Status == "Missing email")} missing · {_preview.Count(r => r.Status == "Invalid email")} invalid · {_preview.Count(r => r.Status == "Duplicate removed - print only")} print-only duplicates";
        if(_previewMailing)SummaryText.Text+=$" - {_preview.Count(r=>r.MailingStatus.Contains("review",StringComparison.OrdinalIgnoreCase))} mailing addresses need review";
        if(_ownerMailingMode)SummaryText.Text=$"{_preview.Count:N0} units - {_preview.Count(r=>r.MailingStatus.StartsWith("Default property"))} property defaults - {_preview.Count(r=>r.NeedsMailingReview)} mailing addresses need review";
        CopyButton.IsEnabled = unique > 0;
        ExportButton.IsEnabled = _preview.Count > 0;
    }
    private void Preview_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
    {
        if (e.EditAction != DataGridEditAction.Commit) return;
        _previewEdited = true;
        Dispatcher.BeginInvoke(new Action(() => { PreviewGrid.CommitEdit(DataGridEditingUnit.Row, true); RefreshSummary(); }));
    }
    private bool CommitPreview() => !_filtersDirty && PreviewGrid.CommitEdit(DataGridEditingUnit.Cell, true) && PreviewGrid.CommitEdit(DataGridEditingUnit.Row, true);
    private void Copy_Click(object sender, RoutedEventArgs e)
    {
        if (!CommitPreview()) return;
        RefreshSummary();
        string addresses = string.Join("; ", _preview.SelectMany(r => GeneratedEmailListService.EmailParts(r.Email)).Where(GeneratedEmailListService.IsValidEmail).Distinct(StringComparer.OrdinalIgnoreCase));
        if (addresses.Length == 0) return;
        try { Clipboard.SetText(addresses); SummaryText.Text += " · Addresses copied."; }
        catch (System.Runtime.InteropServices.ExternalException ex) { ShowError(ex); }
    }
    private void Export_Click(object sender, RoutedEventArgs e)
    {
        if (!CommitPreview()) return;
        RefreshSummary();
        SaveFileDialog picker = new() { Title = "Export recipient list", Filter = "Excel table (*.xlsx)|*.xlsx", DefaultExt = ".xlsx", FileName = _ownerMailingMode?$"Owner-mailing-list-{(DefaultCorporationCombo.SelectedItem as OwnerMailingDefault)?.Corporation}-{DateTime.Now:yyyy-MM-dd}.xlsx":$"General-email-list-{Mode}-{DateTime.Now:yyyy-MM-dd}.xlsx", OverwritePrompt = true };
        if (picker.ShowDialog(OwnerWindow) != true) return;
        if (System.IO.Path.GetFullPath(picker.FileName).Equals(System.IO.Path.GetFullPath(_loadedFile), StringComparison.OrdinalIgnoreCase)) { ShowError(new IOException("Choose another file name. The source backup cannot be overwritten.")); return; }
        try { GeneratedEmailListService.Export(picker.FileName, _preview, _ownerMailingMode||GroupCheck.IsChecked == true, _previewMailing); _previewEdited = false; SummaryText.Text += " · Excel list exported."; }
        catch (Exception ex) when (IsFileError(ex)) { ShowError(ex); }
    }
    private void ManageFolders_Click(object sender,RoutedEventArgs e)
    {
        if(BuildingLink.CorporationConfiguration.Manage(OwnerWindow,ConfirmDiscard))ReloadSavedFolders();
    }

    internal event Action<RecipientFileData,string>? CampaignRequested;
    internal RecipientFileData CampaignSnapshot()
    {
        PreviewGrid.CommitEdit(DataGridEditingUnit.Cell,true);PreviewGrid.CommitEdit(DataGridEditingUnit.Row,true);
        if(_source is null||_filtersDirty||_preview.Count==0)throw new InvalidOperationException("Generate the recipient preview first.");
        // Transfer the displayed preview, including print-only rows left by duplicate removal.
        // Grouping and duplicate removal are independent choices; neither blocks the handoff.
        var data=GeneratedEmailListService.ToRecipientData(_preview,GroupCheck.IsChecked==true);
        string corporation=_shortcuts.FirstOrDefault(f=>SameFolder(f.Path,FolderText.Text))?.Name??string.Join(" ",_source.Rows.Select(r=>GeneratedEmailListService.Value(r,"Building Name")).Distinct());
        foreach(var row in data.Rows)row["Corporation"]=corporation;data.Columns.Add("Corporation");return data;
    }
    private void UseCampaign_Click(object s,RoutedEventArgs e)
    {
        try{CampaignRequested?.Invoke(CampaignSnapshot(),"Generate Email List - "+Path.GetFileName(_loadedFile));}catch(Exception ex){ShowError(ex);}
    }
    private bool SaveShortcuts(List<FolderShortcut> folders)
    {
        try { FolderShortcutStore.Save(folders); _shortcuts = folders; RenderShortcuts(); return true; }
        catch (Exception ex) when (IsFileError(ex)) { ShowError(ex); return false; }
    }
    internal void ReloadSavedFolders()
    {
        try
        {
            var previous=_shortcuts.FirstOrDefault(f=>SameFolder(f.Path,FolderText.Text));
            _shortcuts=FolderShortcutStore.Load();var current=_shortcuts.FirstOrDefault(f=>SameFolder(f.Path,FolderText.Text));
            if(_source is not null&&previous is not null&&current is null)
            {ResetSource();FolderText.Clear();SetBusy(false);SourceText.Text="Corporation folders changed. Select the corporation to load its current backup.";}
            if(_ownerMailingMode&&_source is not null&&previous is not null&&(current is null||previous.UnitPath!=current.UnitPath||previous.Name!=current.Name))
            {_mailingSource=null;_mailingSourcePath="";MarkFiltersChanged();UpdateDefaultCorporation();UpdateMailingControls();OwnerUnitSourceText.Text="Corporation folders changed. Select the corporation again or Refresh to load both backups.";}
            RenderShortcuts();
        }
        catch(Exception ex) when(IsFileError(ex)) { SourceText.Text="Saved folders could not be refreshed. "+ex.Message; }
    }
    private void RenderShortcuts()
    {
        ShortcutButtons.Children.Clear();
        EmptyFoldersText.Visibility = _shortcuts.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        foreach (var folder in _shortcuts) {
            bool selected = SameFolder(folder.Path, FolderText.Text);
            var color = (Color)ColorConverter.ConvertFromString(folder.ButtonColor);
            var tint = Color.FromRgb((byte)(color.R * 0.10 + 255 * 0.90), (byte)(color.G * 0.10 + 255 * 0.90), (byte)(color.B * 0.10 + 255 * 0.90));
            Button button = new() { Content = folder.Name, ToolTip = folder.Path, Padding = new Thickness(14,8,14,8), Margin = new Thickness(0,0,8,5),
                Background = new SolidColorBrush(selected ? color : tint), Foreground = new SolidColorBrush(selected ? Colors.White : color),
                BorderBrush = new SolidColorBrush(color), FontWeight = selected ? FontWeights.SemiBold : FontWeights.Normal };
            button.Click += async (_, _) => { if (ConfirmDiscard()) await SelectFolderAsync(folder.Path); };
            ShortcutButtons.Children.Add(button);
        }

        BackupHeadingText.Text = _shortcuts.FirstOrDefault(f => SameFolder(f.Path, FolderText.Text))?.Name ?? "No corporation selected";
    }
    private static bool SameFolder(string a, string b) => a.TrimEnd('\\','/').Equals(b.TrimEnd('\\','/'), StringComparison.OrdinalIgnoreCase);
    internal static string SuggestFolderName(string path) { var m = Regex.Match(path,@"DEMO[\s_-]*(\d+)",RegexOptions.IgnoreCase); return m.Success ? "DEMO " + m.Groups[1].Value : System.IO.Path.GetFileName(path.TrimEnd('\\','/')); }
    internal static DateTime? BackupDateFromFileName(string path) { var m = Regex.Match(System.IO.Path.GetFileName(path), @"_(\d{4}-\d{2}-\d{2}_\d{6})(?=\.)"); return m.Success && DateTime.TryParseExact(m.Groups[1].Value,"yyyy-MM-dd_HHmmss",CultureInfo.InvariantCulture,DateTimeStyles.None,out var value) ? value : null; }
}
