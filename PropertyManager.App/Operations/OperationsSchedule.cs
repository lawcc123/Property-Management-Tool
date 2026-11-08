using System.Diagnostics;
using System.IO;
using System.Text.Json;

namespace PropertyManager.App.Operations;

internal static class OperationsSchedule
{
    internal static async Task<JsonElement> ExecuteAsync(string action,OperationsSettings settings)
    {
        PropertyManager.App.PublicEdition.BlockLiveOperation();
        var info=new ProcessStartInfo("powershell.exe") {UseShellExecute=false,CreateNoWindow=true,RedirectStandardInput=true,RedirectStandardOutput=true,RedirectStandardError=true};
        foreach(string arg in new[]{"-NoLogo","-NoProfile","-NonInteractive","-ExecutionPolicy","Bypass","-File",Path.Combine(AppContext.BaseDirectory,"Backend","OperationsSchedule.ps1"),"-Action",action}) info.ArgumentList.Add(arg);
        using var process=new Process{StartInfo=info}; process.Start();
        await process.StandardInput.WriteAsync(JsonSerializer.Serialize(new {settings.Time,settings.Days,Executable=Environment.ProcessPath})); process.StandardInput.Close();
        var output=process.StandardOutput.ReadToEndAsync(); var errors=process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync(); string text=await output; string error=await errors;
        if(process.ExitCode!=0) throw new InvalidOperationException("Windows could not "+(action=="Get"?"read":"change")+" the schedule. "+error.Trim());
        using var document=JsonDocument.Parse(text); return document.RootElement.Clone();
    }
}
