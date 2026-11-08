using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using PropertyManager.App.Operations;
using PropertyManager.App.Services;
namespace PropertyManager.App.PersonalizedEmail;
internal sealed class CampaignSmtpTransport:ICampaignTransport
{
    private readonly string? backendDirectory;
    private readonly Func<LoginCredentials?> credentialReader;
    internal CampaignSmtpTransport(string? backendDirectory=null,Func<LoginCredentials?>? credentialReader=null){this.backendDirectory=backendDirectory;this.credentialReader=credentialReader??(()=>WindowsCredentialVault.Read(OperationsRunner.EmailTarget));}
    private Process? process;
    private Task? errors;
    public async Task Connect(PreparedCampaign campaign)
    {
        PropertyManager.App.PublicEdition.BlockLiveOperation();
        var credential=credentialReader()??throw new CampaignNotSubmittedException("Save the Email sending account/app password in Manage credentials first.");
        string backend=backendDirectory??Path.Combine(AppContext.BaseDirectory,"Backend"),script=Path.Combine(backend,"operations_worker.py");bool dev=File.Exists(script);string executable=dev?"python":Path.Combine(backend,"Operations","Operations.Worker.exe");
        var info=new ProcessStartInfo(executable){UseShellExecute=false,CreateNoWindow=true,WorkingDirectory=backend,RedirectStandardInput=true,RedirectStandardOutput=true,RedirectStandardError=true,StandardInputEncoding=new UTF8Encoding(false),StandardOutputEncoding=Encoding.UTF8};info.Environment["PYTHONIOENCODING"]="utf-8";
        if(dev){info.ArgumentList.Add("-u");info.ArgumentList.Add(script);}info.ArgumentList.Add("--action");info.ArgumentList.Add("campaign-delivery");info.ArgumentList.Add("--cancel");info.ArgumentList.Add(Path.Combine(Path.GetTempPath(),"campaign-unused-"+Guid.NewGuid().ToString("N")));
        process=new(){StartInfo=info};process.Start();errors=process.StandardError.ReadToEndAsync();
        await process.StandardInput.WriteLineAsync(JsonSerializer.Serialize(new{campaign,credentials=new{username=credential.Username,password=credential.Password}}));credential=null;
        var ready=await Read();if(ready.GetProperty("type").GetString()!="ready")throw new CampaignNotSubmittedException(ready.TryGetProperty("message",out var m)?m.GetString()??"SMTP connection failed":"SMTP connection failed");
    }
    private async Task<JsonElement> Read()
    {
        string? line=await process!.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(90));
        if(line is null)throw new IOException("The mail component stopped unexpectedly. Check history before retrying.");
        using var json=JsonDocument.Parse(line);return json.RootElement.Clone();
    }
    public async Task Send(DeliveryMessage message,Action beforeSubmission)
    {
        await process!.StandardInput.WriteLineAsync(JsonSerializer.Serialize(new{action="send",id=message.Id}));
        var ready=await Read();if(ready.GetProperty("type").GetString()!="ready_to_submit")throw new CampaignNotSubmittedException(ready.TryGetProperty("message",out var detail)?detail.GetString()??"Message validation failed":"Message validation failed");
        if(ready.GetProperty("id").GetString()!=message.Id)throw new CampaignNotSubmittedException("Mail worker returned an unexpected message identity.");
        beforeSubmission(); // SQLite transaction completes before granting permission to send.
        await process.StandardInput.WriteLineAsync(JsonSerializer.Serialize(new{action="commit",id=message.Id}));
        var result=await Read();
        if(result.GetProperty("type").GetString()=="accepted"&&result.GetProperty("id").GetString()==message.Id)return;
        string error=result.TryGetProperty("message",out var m)?m.GetString()??"SMTP submission failed":"SMTP submission failed";
        if(result.TryGetProperty("uncertain",out var uncertain)&&!uncertain.GetBoolean())throw new CampaignNotSubmittedException(error);
        throw new IOException(error);
    }
    public async ValueTask DisposeAsync()
    {
        if(process is null)return;
        try{if(!process.HasExited){await process.StandardInput.WriteLineAsync("{\"action\":\"quit\"}");await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));}}
        catch{try{if(!process.HasExited)process.Kill(true);}catch{}}
        finally{if(errors is not null){try{await errors;}catch{}}process.Dispose();process=null;}
    }
}
