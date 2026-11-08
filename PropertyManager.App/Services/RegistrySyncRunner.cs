using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;

namespace PropertyManager.App.Services;

/// <summary>Owns one Python worker. Structured stdout and diagnostic stderr are drained concurrently.</summary>
public sealed class RegistrySyncRunner
{
    private Process? process;
    private string? cancelPath;
    private string? retryPath;

    public async Task<int> RunAsync(string workbook, Action<JsonElement> receive,
        Action<string> diagnostic, bool checkOnly = false, LoginCredentials? credentials = null, string? backupFolder = null)
    {
        PropertyManager.App.PublicEdition.BlockLiveOperation();
        if (process is not null) throw new InvalidOperationException("A sync is already running.");
        var source = Path.Combine(AppContext.BaseDirectory, "Backend", "registry_worker.py");
        var development = File.Exists(source);
        var worker = development ? source : Path.Combine(AppContext.BaseDirectory, "Backend", "RegistrySync.Worker.exe");
        if (!File.Exists(worker)) throw new FileNotFoundException("The sync component is missing. Extract the complete app folder and try again.", worker);
        var info = new ProcessStartInfo
        {
            FileName = development ? "python" : worker,
            WorkingDirectory = Path.GetDirectoryName(worker)!,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            RedirectStandardInput = true,
            StandardInputEncoding = new UTF8Encoding(false),

            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8

        };
        info.Environment["PYTHONIOENCODING"] = "utf-8";
        info.Environment["PYTHONUTF8"] = "1";
        if (development)
        {
            foreach (var argument in new[] { "-u", "-X", "utf8", source }) info.ArgumentList.Add(argument);
        }
        foreach (var argument in new[] { "--workbook", workbook })
            info.ArgumentList.Add(argument);
        if(!string.IsNullOrWhiteSpace(backupFolder))
        {
            if(!Path.IsPathFullyQualified(backupFolder))throw new InvalidOperationException("Choose a full registry backup folder path.");
            info.ArgumentList.Add("--backup-folder");info.ArgumentList.Add(backupFolder);
        }
        cancelPath = Path.Combine(Path.GetTempPath(), "PropertyManager-sync-" + Guid.NewGuid().ToString("N") + ".cancel");
        retryPath = cancelPath + ".retry";
        info.ArgumentList.Add("--login-stdin");
        info.ArgumentList.Add("--retry-file");
        info.ArgumentList.Add(retryPath);
        info.ArgumentList.Add("--cancel-file");
        info.ArgumentList.Add(cancelPath);
        if (checkOnly) info.ArgumentList.Add("--check-only");
        using var child = new Process { StartInfo = info };
        try
        {
            if (!child.Start()) throw new InvalidOperationException("The sync component could not start.");
            process = child;
            try
            {
                await child.StandardInput.WriteLineAsync(credentials is null ? "{}" : JsonSerializer.Serialize(credentials));
                await child.StandardInput.FlushAsync();
            }
            catch (IOException) { diagnostic("Sign-in details could not be delivered to the sync component."); }
            finally { child.StandardInput.Close(); }
            credentials = null;
            var stdout = Task.Run(async () =>
            {
                while (await child.StandardOutput.ReadLineAsync() is { } line)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    try
                    {
                        using var document = JsonDocument.Parse(line);
                        receive(document.RootElement.Clone());
                    }
                    catch (JsonException) { diagnostic(line); }
                }
            });
            var stderr = Task.Run(async () =>
            {
                while (await child.StandardError.ReadLineAsync() is { } line) diagnostic(line);
            });
            await child.WaitForExitAsync();
            await Task.WhenAll(stdout, stderr);
            return child.ExitCode;
        }
        finally
        {
            process = null;
            if (cancelPath is not null)
            {
                try { File.Delete(cancelPath); } catch (IOException) { }
                cancelPath = null;
            }
            if (retryPath is not null)
            {
                try { File.Delete(retryPath); } catch (IOException) { }
                retryPath = null;
            }
        }
    }

    public async Task RetryLoginAsync()
    {
        var path = retryPath;
        if (process is null || process.HasExited || path is null) return;
        await File.WriteAllTextAsync(path, Guid.NewGuid().ToString("N"));
    }

    public async Task CancelAsync()
    {
        var child = process;
        if (child is null || child.HasExited || cancelPath is null) return;
        try
        {
            await File.WriteAllTextAsync(cancelPath, "cancel");
        }
        catch (IOException) { /* Worker exited before the request arrived. */ }
        catch (InvalidOperationException) { /* Process already exited. */ }
    }
}

public sealed class AppSettings
{
    public string Workbook { get; set; } = "";
    public string LastSync { get; set; } = "";
    public string RegistryBackupFolder { get; set; } = "";
    public string EffectiveRegistryBackupFolder()=>!string.IsNullOrWhiteSpace(RegistryBackupFolder)?RegistryBackupFolder:
        string.IsNullOrWhiteSpace(Workbook)?"Registry Backups beside the selected workbook":Path.Combine(Path.GetDirectoryName(Path.GetFullPath(Workbook))!,"Registry Backups");
    public static string DataDirectory => Environment.GetEnvironmentVariable("PROPERTYMANAGER_DEMO_HOME") is { Length: > 0 } diagnostic
        ? Path.GetFullPath(diagnostic)
        : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PropertyManagerDemo");
    private static string SettingsPath => Path.Combine(DataDirectory, "settings.json");

    public static AppSettings Load()
    {
        try { return JsonSerializer.Deserialize<AppSettings>(Settings.SettingsDatabase.Read(SettingsPath)??"{}") ?? new(); }
        catch (IOException) { return new(); }
        catch (UnauthorizedAccessException) { return new(); }
        catch (JsonException) { return new(); }
    }

    public void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
        Settings.SettingsDatabase.Write(SettingsPath, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
    }
}

