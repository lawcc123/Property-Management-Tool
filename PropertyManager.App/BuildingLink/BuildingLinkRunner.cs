using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using PropertyManager.App.Services;

namespace PropertyManager.App.BuildingLink;

internal sealed class BuildingLinkRunner
{
    private readonly Func<LoginCredentials?> credentialReader;
    private readonly string? backendDirectory;
    internal BuildingLinkRunner(Func<LoginCredentials?>? credentialReader=null,string? backendDirectory=null)
    {
        this.credentialReader=credentialReader??(()=>WindowsCredentialVault.Read(BuildingLinkCredentialStore.Target));
        this.backendDirectory=backendDirectory;
    }
    private string? cancelPath;
    private bool busy;
    internal bool IsRunning => busy;
    internal void Cancel() { if(cancelPath is not null) File.WriteAllText(cancelPath,""); }
    internal async Task<int> RunAsync(Action<JsonElement> receive)
    {
        PropertyManager.App.PublicEdition.BlockLiveOperation();
        if(busy) throw new InvalidOperationException("A BuildingLink download is already running.");
        Directory.CreateDirectory(AppSettings.DataDirectory);
        FileStream runLock;
        try { runLock=new FileStream(Path.Combine(AppSettings.DataDirectory,"buildinglink-run.lock"),FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None); }
        catch(IOException) { throw new InvalidOperationException("Another BuildingLink download is already running. Wait for it to finish."); }
        using var heldLock=runLock;
        busy=true;
        try
        {
            var settings=BuildingLinkSettings.Load(); settings.Validate();settings.Save();
            var login=credentialReader()
                ?? throw new InvalidOperationException("Save a BuildingLink login in Manage credentials first.");
            string backend=backendDirectory??Path.Combine(AppContext.BaseDirectory,"Backend");
            bool development=File.Exists(Path.Combine(backend,"buildinglink_worker.py"));
            string worker=Path.Combine(backend,development?"buildinglink_worker.py":"BuildingLink/BuildingLink.Worker.exe");
            if(!File.Exists(worker)) throw new FileNotFoundException("The BuildingLink download component is missing.");
            Directory.CreateDirectory(AppSettings.DataDirectory);
            cancelPath=Path.Combine(AppSettings.DataDirectory,"buildinglink-cancel-"+Guid.NewGuid().ToString("N"));
            var info=new ProcessStartInfo(development?"python":worker) { WorkingDirectory=backend,UseShellExecute=false,CreateNoWindow=true,
                RedirectStandardInput=true,RedirectStandardOutput=true,RedirectStandardError=true,StandardInputEncoding=new UTF8Encoding(false),StandardOutputEncoding=Encoding.UTF8,StandardErrorEncoding=Encoding.UTF8 };
            info.Environment["PYTHONIOENCODING"]="utf-8";
            if(development) { info.ArgumentList.Add("-u"); info.ArgumentList.Add(worker); }
            foreach(string value in new[]{"--settings",BuildingLinkSettings.SettingsPath,"--result",BuildingLinkSettings.ResultPath,"--cancel",cancelPath}) info.ArgumentList.Add(value);
            using var process=new Process{StartInfo=info};
            if(!process.Start()) throw new InvalidOperationException("Could not start BuildingLink downloader.");
            await process.StandardInput.WriteLineAsync(JsonSerializer.Serialize(new { username=login.Username,password=login.Password }));
            process.StandardInput.Close(); login=null;
            // Drain stderr but never retain raw automation errors or form data in the app log.
            var errors=Task.Run(async()=> { while(await process.StandardError.ReadLineAsync() is not null) { } });
            bool completed=false;
            while(await process.StandardOutput.ReadLineAsync() is { } line)
            {
                try { using var data=JsonDocument.Parse(line); if(data.RootElement.TryGetProperty("type",out var type)&&type.GetString()=="complete")completed=true;receive(data.RootElement.Clone()); }
                catch(JsonException) { }
            }
            await process.WaitForExitAsync(); await errors;
            if(!completed)
                BuildingLinkSettings.AtomicWrite(BuildingLinkSettings.ResultPath,JsonSerializer.Serialize(new {success=false,message="The downloader stopped before completion. Check Microsoft Edge and the download components.",finished=DateTimeOffset.Now,files=Array.Empty<string>()}));
            if(completed && process.ExitCode==0 && File.Exists(BuildingLinkSettings.ResultPath))
                BuildingLinkSettings.AtomicWrite(BuildingLinkSettings.SuccessPath,File.ReadAllText(BuildingLinkSettings.ResultPath));
            return process.ExitCode==0&&!completed?1:process.ExitCode;
        }
        catch(Exception ex)
        {
            Directory.CreateDirectory(AppSettings.DataDirectory);
            BuildingLinkSettings.AtomicWrite(BuildingLinkSettings.ResultPath,JsonSerializer.Serialize(new { success=false,message=ex.Message,finished=DateTimeOffset.Now }));
            throw;
        }
        finally
        {
            try { BuildingLinkSettings.AtomicWrite(BuildingLinkSettings.ProgressPath,JsonSerializer.Serialize(new {running=false,updated=DateTimeOffset.UtcNow})); }
            catch(Exception ex) when(ex is IOException or UnauthorizedAccessException) { }
            try { if(cancelPath is not null && File.Exists(cancelPath)) File.Delete(cancelPath); }
            catch(Exception ex) when(ex is IOException or UnauthorizedAccessException) { }
            cancelPath=null; busy=false; 
        }
    }
}
