using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using PropertyManager.App.Services;
using Microsoft.Win32;

namespace PropertyManager.App;

public partial class MainWindow : Window
{
    private readonly ObservableCollection<SyncRecord> records = new();
    private readonly RegistrySyncRunner runner = new();
    private readonly AppSettings settings = AppSettings.Load();
    private bool running;
    private bool saving;
    private string? backupPath;
    private JsonElement? completed;
    private string lastError = "";
    private bool cancelled;
    private readonly Stopwatch elapsed = new();
    private readonly DispatcherTimer clock = new() { Interval = TimeSpan.FromSeconds(1) };
    private int reviewed;
    private int totalRows;
    private string loginOrigin = "";
    private string? savedLoginUsername;
    private bool waitingForLogin;
    private LoginCredentials? activeCredentials;

    public MainWindow() : this(null) { }

    internal MainWindow(string? credentialOrigin)
    {
        InitializeComponent();InitializeAppearance();
        var version = typeof(MainWindow).Assembly.GetName().Version;
        AppVersionText.Text = $"Version {version?.ToString(3)} \u00B7 Workspace";
        WorkbookIcon.Source = WindowsFileIcon.Load(".xlsx");
        ResultsGrid.ItemsSource = records;
        InitializeLogin(credentialOrigin);
        clock.Tick += (_, _) => UpdateProgressText();
        if (File.Exists(settings.Workbook)) SelectWorkbook(settings.Workbook);
        UpdateBackupFolder();
        LastSync.Text = string.IsNullOrEmpty(settings.LastSync) ? "No completed sync yet" : "Last sync: " + settings.LastSync;
        Closing += OnClosing;
        Loaded += async (_, _) =>
        {
            var args = Environment.GetCommandLineArgs();
            if(args.Contains("--personalized-email"))ShowPage(7);
            var validateIndex = Array.IndexOf(args, "--validate-workbook");
            var reportIndex = Array.IndexOf(args, "--report");
            if (validateIndex >= 0 && validateIndex + 1 < args.Length && reportIndex >= 0 && reportIndex + 1 < args.Length)
            {
                SelectWorkbook(args[validateIndex + 1]);
                await ExecuteSyncAsync(true);
                await File.WriteAllTextAsync(args[reportIndex + 1], JsonSerializer.Serialize(new
                {
                    success = completed is not null,
                    developerTools = BuildFeatures.DeveloperToolsEnabled,
                    error = lastError,
                    completion = completed,
                    rowsDisplayed = records.Count
                }));
                Close();
                return;
            }
            var captureIndex = Array.IndexOf(args, "--capture");
            if (captureIndex >= 0 && captureIndex + 1 < args.Length)
            {
                await Task.Delay(300);
                UpdateLayout();
                var bitmap = new RenderTargetBitmap((int)ActualWidth, (int)ActualHeight, 96, 96, PixelFormats.Pbgra32);
                bitmap.Render(this);
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmap));
                using (var stream = File.Create(args[captureIndex + 1])) encoder.Save(stream);
                Close();
            }
        };
    }

    private void InitializeLogin(string? credentialOrigin)
    {
        try
        {
            loginOrigin = credentialOrigin ?? YardiCredentialStore.ConfiguredOrigin();
            var savedUsername = YardiCredentialStore.Username(loginOrigin);
            if (savedUsername is not null)
            {
                savedLoginUsername = savedUsername;
                LoginUsername.Text = savedUsername;
                UseAutofill.IsChecked = RememberLogin.IsChecked = true;
                ForgetLoginButton.IsEnabled = true;
                AutofillStatus.Text = "Saved login available for this Windows account. The masked Saved indicator means your password is ready to reuse. Type here to replace it.";
            }
            else AutofillStatus.Text = "Optional autofill for " + new Uri(loginOrigin).Host + ". Login is submitted once; complete verification yourself. Elevate is selected automatically.";
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or UriFormatException or InvalidOperationException or Win32Exception)
        { AutofillStatus.Text = "Optional saved sign-in is unavailable. You can still sign in manually in the browser."; }
    }

    private void Autofill_Changed(object sender, RoutedEventArgs e)
    {
        if (LoginFields is not null) LoginFields.IsEnabled = UseAutofill.IsChecked == true;
        UpdateSavedPasswordIndicator();
    }

    private void SavedPassword_Changed(object sender, RoutedEventArgs e) => UpdateSavedPasswordIndicator();
    private void UpdateSavedPasswordIndicator()
    {
        if (SavedPasswordIndicator is null || LoginPassword is null || LoginUsername is null) return;
        SavedPasswordIndicator.Visibility = savedLoginUsername is not null &&
            string.Equals(savedLoginUsername, LoginUsername.Text.Trim(), StringComparison.Ordinal) &&
            LoginPassword.Password.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private LoginCredentials? PrepareLogin()
    {
        if (UseAutofill.IsChecked != true) return null;
        if (string.IsNullOrEmpty(loginOrigin)) throw new InvalidOperationException("Optional autofill requires a configured HTTPS Yardi login address.");
        var username = LoginUsername.Text.Trim();
        if (string.IsNullOrWhiteSpace(username)) throw new InvalidOperationException("Enter your Yardi username or email, or turn autofill off.");
        var password = LoginPassword.Password;
        if (string.IsNullOrEmpty(password))
        {
            var saved = YardiCredentialStore.Read(loginOrigin);
            if (saved is null || !string.Equals(saved.Username, username, StringComparison.Ordinal))
                throw new InvalidOperationException("Enter a password for this username, or turn autofill off to sign in manually.");
            password = saved.Password;
        }
        if (RememberLogin.IsChecked != true)
        {
            YardiCredentialStore.Delete(loginOrigin);
            savedLoginUsername = null;
            UpdateSavedPasswordIndicator();
            if (string.IsNullOrEmpty(LoginPassword.Password)) LoginPassword.Password = password;
            ForgetLoginButton.IsEnabled = false;
        }
        return new LoginCredentials { Origin = loginOrigin, Username = username, Password = password };
    }

    private void ForgetLogin_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            YardiCredentialStore.Delete(loginOrigin);
            savedLoginUsername = null;
            UpdateSavedPasswordIndicator();
            LoginPassword.Clear();
            RememberLogin.IsChecked = UseAutofill.IsChecked = false;
            ForgetLoginButton.IsEnabled = false;
            AutofillStatus.Text = "Saved login removed from Windows Credential Manager. Manual sign-in is still available.";
        }
        catch (Win32Exception) { AutofillStatus.Text = "Windows could not remove the saved login. Please try again."; }
    }

    private void SelectWorkbook(string path)
    {
        settings.Workbook = Path.GetFullPath(path);
        WorkbookName.Text = Path.GetFileName(path);
        WorkbookIcon.Source = WindowsFileIcon.Load(Path.GetExtension(path));
        WorkbookPath.Text = Path.GetDirectoryName(settings.Workbook);
        WorkbookPath.ToolTip = settings.Workbook;
        SyncButton.IsEnabled = true;
        OpenRegistryButton.IsEnabled = !running;
        UpdateBackupFolder();
    }

    private void OpenRegistry_Click(object sender, RoutedEventArgs e)
    {
        if(running)return;
        try{Process.Start(WorkbookFile.Launch(settings.Workbook,true));}
        catch(Exception ex){MessageBox.Show(this,ex.Message,"Open registry workbook",MessageBoxButton.OK,MessageBoxImage.Warning);}
    }

    private void Browse_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Choose your cheque registry workbook",
            Filter = "Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
            CheckFileExists = true,
            Multiselect = false
        };
        if (dialog.ShowDialog(this) != true) return;
        SelectWorkbook(dialog.FileName);
        SaveSettings();
        Log("Workbook selected: " + dialog.FileName);
        StageText.Text = "Workbook selected. Close it in Excel, then choose Start sync.";
    }

    private async void Sync_Click(object sender, RoutedEventArgs e)
    {
        if (running)
        {
            if (!waitingForLogin) return;
            try
            {
                await runner.RetryLoginAsync();
                StageText.Text = "Retry requested. The browser will reload the Yardi sign-in page.";
            }
            catch (IOException) { StageText.Text = "The sign-in retry could not be requested. Finish in the browser or cancel and start again."; }
            return;
        }
        await ExecuteSyncAsync();
    }

    private async Task ExecuteSyncAsync(bool checkOnly = false)
    {
        if (running) return;
        if (!File.Exists(settings.Workbook))
        {
            StageText.Text = "Choose an existing registry workbook first.";
            return;
        }
        try { activeCredentials = checkOnly ? null : PrepareLogin(); }
        catch (Exception ex) when (ex is InvalidOperationException or Win32Exception or IOException)
        {
            LoginPanel.IsExpanded = true;
            AutofillStatus.Text = ex.Message;
            StageText.Text = "Check the optional sign-in details, or turn autofill off to sign in manually.";
            return;
        }
        waitingForLogin = false;
        LoginPanel.IsEnabled = RegistryBackupFolderControls.IsEnabled = false;
        running = true;
        saving = false;
        cancelled = false;
        completed = null;
        backupPath = null;
        lastError = "";

        SaveSettings();
        records.Clear();
        AttentionPanel.Visibility = ResultsGrid.Visibility = NoAttentionText.Visibility = Visibility.Collapsed;
        reviewed = totalRows = 0;
        elapsed.Restart();
        clock.Start();
        UpdateProgressText();
        BackupButton.Visibility = Visibility.Collapsed;
        ReviewedCount.Text = UpdatedCount.Text = AttentionCount.Text = "0";
        BrowseButton.IsEnabled = OpenRegistryButton.IsEnabled = SyncButton.IsEnabled = false;
        CancelButton.Visibility = Visibility.Visible;
        CancelButton.IsEnabled = true;
        SyncButton.Content = "Sync running…";
        RunStatus.Text = checkOnly ? "Checking workbook" : "Starting sync";
        ConnectionStatus.Text = "Starting…";
        try { File.WriteAllText(LogPath, "Registry Sync started." + Environment.NewLine); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        RunProgress.Value = 0;
        RunProgress.IsIndeterminate = true;
        try
        {
            var exitCode = await runner.RunAsync(settings.Workbook,
                data => Dispatcher.Invoke(() => HandleEvent(data)),
                message => Dispatcher.Invoke(() => Log(message)), checkOnly, activeCredentials, settings.RegistryBackupFolder);
            if (completed is { } final && exitCode == 0)
            {
                RunProgress.IsIndeterminate = false;
                RunProgress.Value = 100;
                var updated = Number(final, "updated");
                var attention = Number(final, "attention");
                var skipped = Number(final, "skipped");
                ReviewedCount.Text = Number(final, "reviewed").ToString();
                UpdatedCount.Text = updated.ToString();
                AttentionCount.Text = attention.ToString();
                RunStatus.Text = checkOnly ? "Workbook verified" : attention > 0 ? "Complete · review flagged rows" : "Sync complete";
                ConnectionStatus.Text = "Ready for the next sync";
                backupPath = Text(final, "backup");
                BackupButton.Visibility = string.IsNullOrEmpty(backupPath) ? Visibility.Collapsed : Visibility.Visible;
                StageText.Text = checkOnly
                    ? $"Workbook verified: {Number(final, "total")} invoice rows. No browser opened and no files changed."
                    : $"{Number(final, "reviewed")} reviewed · {updated} updated · {attention} need attention. " +
                      (skipped > 0 ? $"{skipped} intentionally excluded. " : "") +
                      (updated > 0 ? "Workbook saved; an original backup is available." : "No cell changes; workbook was not saved.");
                if (!checkOnly)
                {
                    settings.LastSync = DateTime.Now.ToString("g");
                    LastSync.Text = "Last sync: " + settings.LastSync;
                    SaveSettings();
                }
                if (!checkOnly) ShowAttentionResults(true);
                Log(StageText.Text);
                if (!string.IsNullOrEmpty(backupPath)) Log("Original backup: " + backupPath);
            }
            else if (cancelled || exitCode == 2)
            {
                RunStatus.Text = "Cancelled";
                ConnectionStatus.Text = "Sync cancelled";
                StageText.Text = "Sync cancelled. The selected workbook was not changed.";
                MarkUnsaved();
            }
            else
            {
                throw new InvalidOperationException(string.IsNullOrEmpty(lastError)
                    ? $"The sync component exited with code {exitCode}. Re-extract the complete app folder and try again."
                    : lastError);
            }
        }
        catch (Exception ex)
        {
            lastError = ex.Message;
            RunStatus.Text = "Sync could not complete";
            ConnectionStatus.Text = "Sync could not complete";
            StageText.Text = ex is Win32Exception
                ? "Could not launch the sync component. This development build requires Python; the release will include its own runtime."
                : ex.Message;
            Log("ERROR: " + ex.Message);
            MarkUnsaved();
            if (records.Count > 0) ShowAttentionResults(false);
        }
        finally
        {
            RunProgress.IsIndeterminate = false;
            elapsed.Stop();
            clock.Stop();
            UpdateProgressText();
            running = saving = waitingForLogin = false;
            activeCredentials = null;
            LoginPanel.IsEnabled = RegistryBackupFolderControls.IsEnabled = true;
            UpdateBackupFolder();
            BrowseButton.IsEnabled = OpenRegistryButton.IsEnabled = SyncButton.IsEnabled = true;
            SyncButton.Content = "Start sync  →";
            CancelButton.IsEnabled = false;
        }
    }

    private void HandleEvent(JsonElement data)
    {
        switch (Text(data, "type"))
        {
            case "stage":
                StageText.Text = Text(data, "message");
                var stage = Text(data, "stage");
                waitingForLogin = stage == "sign_in";
                SyncButton.IsEnabled = waitingForLogin;
                SyncButton.Content = waitingForLogin ? "Retry sign-in" : "Sync running...";
                if (!cancelled) RunStatus.Text = stage switch
                {
                    "session_check" => "Checking saved Yardi session",
                    "sign_in" => "Waiting for Yardi sign-in",
                    "lookup" => "Loading P2P records",
                    "syncing" => "Reviewing registry",
                    "saving" => "Saving workbook",
                    _ => "Reading workbook"
                };
                ConnectionStatus.Text = stage switch
                {
                    "session_check" => "Checking saved session",
                    "sign_in" => "Complete verification in the browser",
                    "lookup" or "syncing" or "saving" => "Yardi session connected",
                    _ => "Preparing connection"
                };
                if (stage == "saving")
                {
                    saving = true;
                    CancelButton.IsEnabled = false;
                    RunProgress.IsIndeterminate = true;
                }
                break;
            case "login_state":
                var state = Text(data, "state");
                waitingForLogin = state != "connected";
                SyncButton.IsEnabled = waitingForLogin;
                SyncButton.Content = waitingForLogin ? "Retry sign-in" : "Sync running...";
                StageText.Text = Text(data, "message");
                RunStatus.Text = state switch
                {
                    "connected" => "P2P session confirmed",
                    "submitted" => "Waiting for verification",
                    "elevate" => "Opening Elevate (Production)",
                    "handoff_wait" => "Confirming P2P access",
                    _ => "Waiting for Yardi sign-in"
                };
                ConnectionStatus.Text = waitingForLogin ? "Finish sign-in in the browser" : "Yardi session connected";
                if (state == "connected")
                {
                    if (activeCredentials is not null && RememberLogin.IsChecked == true)
                    {
                        try
                        {
                            // Remember is explicit user consent. A cached session or manual
                            // sign-in can confirm access without any browser autofill event.
                            YardiCredentialStore.Save(activeCredentials);
                            Log("Yardi login remembered securely for this Windows account.");
                            savedLoginUsername = activeCredentials.Username;
                            LoginPassword.Clear();
                            UpdateSavedPasswordIndicator();
                            ForgetLoginButton.IsEnabled = true;
                            AutofillStatus.Text = "Saved in Windows Credential Manager for this Windows account. The masked Saved indicator means your password is ready to reuse.";
                        }
                        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException)
                        {
                            AutofillStatus.Text = "Signed in, but Windows could not remember the login. You can still continue this sync.";
                            Log("Yardi login could not be remembered in Windows Credential Manager.");
                        }
                    }
                    activeCredentials = null;
                }
                break;
            case "validated":
                totalRows = Number(data, "total");
                UpdateProgressText();
                Log($"Workbook checked: {Number(data, "total")} invoice rows across {Number(data, "sheets")} sheets; {Number(data, "excluded")} intentionally excluded.");
                break;
            case "row":
                var reason = Text(data, "reason");
                if (!string.IsNullOrWhiteSpace(reason))
                    records.Add(new SyncRecord($"{Text(data, "sheet")}\nRow {Number(data, "row")}", Text(data, "vendor"),
                        Text(data, "invoice"), Text(data, "po"), Text(data, "ir"), reason));
                reviewed = Number(data, "reviewed");
                totalRows = Number(data, "total");
                UpdateProgressText();
                ReviewedCount.Text = Number(data, "reviewed").ToString();
                UpdatedCount.Text = Number(data, "updated").ToString();
                AttentionCount.Text = Number(data, "attention").ToString();
                RunProgress.IsIndeterminate = false;
                RunProgress.Value = 95.0 * Number(data, "reviewed") / Math.Max(1, Number(data, "total"));
                UpdateProgressText();
                break;
            case "log": Log(Text(data, "message")); break;
            case "complete": completed = data.Clone(); break;
            case "error": lastError = Text(data, "message"); Log("ERROR: " + lastError); break;
            case "cancelled": cancelled = true; Log(Text(data, "message")); break;
        }
    }

    private void MarkUnsaved()
    {
        UpdatedCount.Text = "0";
    }

    private async void Cancel_Click(object sender, RoutedEventArgs e) => await CancelRunAsync();

    private async Task CancelRunAsync()
    {
        if (!running || saving) return;
        CancelButton.IsEnabled = false;
        RunStatus.Text = "Cancellation requested";
        StageText.Text = "Stopping after the current lookup. The workbook will not be saved.";
        await runner.CancelAsync();
    }

    private void OnClosing(object? sender, CancelEventArgs e)
    {
        if(personalizedEmailView?.IsBusy==true){e.Cancel=true;personalizedEmailView.Cancel();return;}
        if(operationsView?.IsBusy==true){e.Cancel=true;operationsView.Cancel();return;}
        if(outlookSyncView?.IsBusy==true){e.Cancel=true;outlookSyncView.Cancel();return;}
        if(envelopeView?.IsBusy==true){e.Cancel=true;envelopeView.Cancel();return;}
        if(buildingLinkView?.IsRunning==true){e.Cancel=true;buildingLinkView.Cancel();return;}
        if(buildingLinkView?.ScheduleBusy==true){e.Cancel=true;return;}
        if(personalizedEmailView is not null&&!personalizedEmailView.ConfirmLeave()){e.Cancel=true;return;}
        if (!running) { if (emailListView is not null && !emailListView.ConfirmDiscard()) e.Cancel = true; if(!e.Cancel&&ownerMailingView is not null&&!ownerMailingView.ConfirmDiscard())e.Cancel=true; return; }
        e.Cancel = true;
        if (saving) StageText.Text = "Please wait for the backup and save to finish before closing.";
        else _ = CancelRunAsync();
    }

    private void Backup_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(backupPath)) return;
        try { Process.Start(new ProcessStartInfo("explorer.exe") { Arguments = $"/select,\"{backupPath}\"", UseShellExecute = true }); }
        catch (Exception ex) { Log("Could not open backup folder: " + ex.Message); }
    }

    private void ShowAttentionResults(bool savedRun)
    {
        AttentionPanel.Visibility = Visibility.Visible;
        AttentionTitle.Text = savedRun ? $"Rows needing attention ({records.Count})" : "Rows flagged before sync stopped";
        AttentionDescription.Text = savedRun
            ? "Only flagged rows are shown. Use the sheet and Excel row number to find each record."
            : "The sync did not finish saving. These are review flags from the rows processed so far.";
        ResultsGrid.Visibility = records.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        NoAttentionText.Visibility = records.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        Dispatcher.BeginInvoke(new Action(() => AttentionPanel.BringIntoView()));
    }

    private void UpdateProgressText()
    {
        RegistryNavStatus.Visibility = running ? Visibility.Visible : Visibility.Collapsed;
        RegistryNavStatus.Text = totalRows > 0 ? $"{reviewed} / {totalRows} reviewed" : "Sync in progress";
        var duration = elapsed.Elapsed.ToString(@"hh\:mm\:ss");
        ProgressMeta.Text = totalRows > 0
            ? $"{Math.Round(RunProgress.Value)}% · {reviewed} / {totalRows} rows · {duration} elapsed"
            : $"{duration} elapsed";
    }

    private static string LogPath
    {
        get
        {
            var directory = AppSettings.DataDirectory;
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, "last-sync.log");
        }
    }

    private void Log(string message)
    {
        try { File.AppendAllText(LogPath, message + Environment.NewLine); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }

    private void SaveSettings()
    {
        try { settings.Save(); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        { Log("Could not save app preferences: " + ex.Message); }
    }

    private static string Text(JsonElement data, string property) =>
        data.TryGetProperty(property, out var value) && value.ValueKind != JsonValueKind.Null ? value.ToString() : "";
    private static int Number(JsonElement data, string property) =>
        data.TryGetProperty(property, out var value) && value.TryGetInt32(out var number) ? number : 0;
}

public record SyncRecord(string Location, string Vendor, string Invoice, string PO, string IR, string Reason);

