using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using PropertyManager.App.Services;

namespace PropertyManager.App.Operations;

internal sealed class OperationsRunner
{
    internal const string EmailTarget="PropertyManagerDemo/OperationsReminderEmail";
    private readonly Func<string,LoginCredentials?> credentialReader;
    private readonly string? backendDirectory;
    internal OperationsRunner(Func<string,LoginCredentials?>? credentialReader=null,string? backendDirectory=null)
    {this.credentialReader=credentialReader??WindowsCredentialVault.Read;this.backendDirectory=backendDirectory;}
    private string? cancelPath;
    internal bool IsRunning {get;private set;}
    internal void Cancel(){if(cancelPath is not null)File.WriteAllText(cancelPath,"");}
    internal async Task<JsonElement> RunAsync(string action,OperationsSettings settings,JsonElement? preview,Action<JsonElement> receive,bool scheduled=false)
    {
        PropertyManager.App.PublicEdition.BlockLiveOperation();
        if(IsRunning)throw new InvalidOperationException("A reminder operation is already running.");
        Directory.CreateDirectory(OperationsSettings.DirectoryPath);
        FileStream runLock;
        try{runLock=new(Path.Combine(OperationsSettings.DirectoryPath,"run.lock"),FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);}catch(IOException){throw new InvalidOperationException("Another reminder operation is running. Wait for it to finish.");}
        using var held=runLock;IsRunning=true;OperationsHistory? history=null;string? run=null;bool submitting=false;string lastStage="Starting";JsonElement? workerFailure=null;
        try
        {
            history=new OperationsHistory();settings.Validate(action=="send");
            if(action=="send"&&!preview.HasValue)throw new InvalidOperationException("Refresh and review the reminder before sending.");
            string? token=action=="send"?preview?.GetProperty("id").GetString():null;
            string? scheduleToken=scheduled?TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow,TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")).ToString("yyyy-MM-dd"):null;
            run=history.Start(scheduled?"Scheduled preview":action,token,scheduleToken);
            cancelPath=Path.Combine(OperationsSettings.DirectoryPath,"cancel-"+Guid.NewGuid().ToString("N"));
            var result=await WorkerAsync(action,settings,preview,item=>
            {
                if(item.TryGetProperty("message",out var stage))lastStage=stage.GetString()??lastStage;
                if(item.TryGetProperty("submitting",out var sent)&&sent.GetBoolean()){submitting=true;history.Update(run,"Submitting - delivery uncertain","SMTP submission started. Do not automatically resend if this run stops.");}
                receive(item);
            });
            if(!result.GetProperty("success").GetBoolean())
            {
                workerFailure=result;
                string message=result.GetProperty("message").GetString()??"Reminder operation failed.";
                if(result.TryGetProperty("cancelled",out var cancelled)&&cancelled.GetBoolean())throw new OperationCanceledException(message);
                history.Update(run,submitting?"Failed - delivery uncertain":result.TryGetProperty("cancelled",out var c)&&c.GetBoolean()?"Cancelled":"Failed",result.GetRawText());
                throw new InvalidOperationException(message);
            }
            var data=result.GetProperty("result").Clone();
            if(scheduled)
            {
                if(File.Exists(cancelPath))throw new OperationCanceledException();
                settings.Validate(true);
                history.Update(run,"Scheduled send preparing",data.GetRawText());
                var sent=await WorkerAsync("send",settings,data,item=>{if(item.TryGetProperty("message",out var stage))lastStage=stage.GetString()??lastStage;if(item.TryGetProperty("submitting",out var v)&&v.GetBoolean()){submitting=true;history.Update(run,"Submitting - delivery uncertain",data.GetRawText());}receive(item);});
                if(!sent.GetProperty("success").GetBoolean()){workerFailure=sent;throw new InvalidOperationException(sent.GetProperty("message").GetString());}
                history.Update(run,"Accepted by mail server",JsonSerializer.Serialize(new{settings,preview=data,result=sent.GetProperty("result")}));return sent.GetProperty("result").Clone();
            }
            history.Update(run,action=="send"?"Accepted by mail server":data.GetProperty("warnings").GetArrayLength()>0?"Preview ready - warnings":"Preview ready",JsonSerializer.Serialize(new{settings,preview,result=data}));return data;
        }
        catch(Exception ex)
        {
            if(run is not null)history!.Update(run,submitting?"Failed - delivery uncertain":ex is OperationCanceledException?"Cancelled":"Failed",JsonSerializer.Serialize(new{settings,preview,lastStep=lastStage,message=ex.Message,workerFailure}));
            throw;
        }
        finally{if(cancelPath is not null){try{File.Delete(cancelPath);}catch(IOException){}}cancelPath=null;IsRunning=false;}
    }
    private async Task<JsonElement> WorkerAsync(string action,OperationsSettings settings,JsonElement? preview,Action<JsonElement> receive)
    {
        string backend=backendDirectory??Path.Combine(AppContext.BaseDirectory,"Backend");bool development=File.Exists(Path.Combine(backend,"operations_worker.py"));string worker=Path.Combine(backend,development?"operations_worker.py":"Operations/Operations.Worker.exe");
        if(!File.Exists(worker))throw new FileNotFoundException("The reservations reminder component is missing.");
        LoginCredentials? login=action=="preview"?credentialReader(BuildingLinkCredentialStore.Target):null;
        LoginCredentials? email=action=="send"?credentialReader(EmailTarget):null;
        var info=new ProcessStartInfo(development?"python":worker){UseShellExecute=false,CreateNoWindow=true,WorkingDirectory=backend,RedirectStandardInput=true,RedirectStandardOutput=true,RedirectStandardError=true,StandardInputEncoding=new UTF8Encoding(false),StandardOutputEncoding=Encoding.UTF8,StandardErrorEncoding=Encoding.UTF8};info.Environment["PYTHONIOENCODING"]="utf-8";
        if(development){info.ArgumentList.Add("-u");info.ArgumentList.Add(worker);}foreach(string arg in new[]{"--action",action,"--cancel",cancelPath!})info.ArgumentList.Add(arg);
        using var process=new Process{StartInfo=info};process.Start();
        await process.StandardInput.WriteLineAsync(JsonSerializer.Serialize(new{settings,preview,credentials=new{buildinglink=new{username=login?.Username??"",password=login?.Password??""},email=new{username=email?.Username??"",password=email?.Password??""}}}));process.StandardInput.Close();login=email=null;
        var errors=Task.Run(async()=>{while(await process.StandardError.ReadLineAsync() is not null){}});
        JsonElement? complete=null;
        while(await process.StandardOutput.ReadLineAsync() is { } line)
        {
            try{using var document=JsonDocument.Parse(line);var item=document.RootElement.Clone();if(item.GetProperty("type").GetString()=="complete")complete=item;else receive(item);}catch(JsonException){}
        }
        await process.WaitForExitAsync();await errors;
        return complete??throw new InvalidOperationException("The reminder worker stopped unexpectedly. Check run history before retrying an email submission.");
    }
}
