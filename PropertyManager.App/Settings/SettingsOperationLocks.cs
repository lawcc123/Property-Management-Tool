using System.IO;
using PropertyManager.App.Services;
namespace PropertyManager.App.Settings;
internal sealed class SettingsOperationLocks:IDisposable
{
    private readonly List<FileStream> held=[];
    internal SettingsOperationLocks(bool settings=true)
    {
        var root=SettingsDatabase.Root;
        var paths=new List<string>{Path.Combine(root,"buildinglink-run.lock"),Path.Combine(root,"OperationsReminder","run.lock"),Path.Combine(root,"OutlookSync","sync.lock"),Path.Combine(root,"PersonalizedEmail","Sending","sending.lock"),Path.Combine(root,"EnvelopePrinter","history.lock")};
        if(string.IsNullOrEmpty(Environment.GetEnvironmentVariable("PROPERTYMANAGER_DEMO_HOME")))paths.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"PropertyManagerDemo_P2P","desktop-sync.lock"));
        if(settings)paths.Add(Path.Combine(root,"settings-write.lock"));
        try{foreach(string path in paths)if(Directory.Exists(Path.GetDirectoryName(path)))held.Add(new FileStream(path,FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None));}
        catch{Dispose();throw new IOException("Another operation is active. Finish sync, downloading, sending and printing before storage maintenance.");}
    }
    public void Dispose(){foreach(var file in held.AsEnumerable().Reverse())file.Dispose();held.Clear();}
}
