using System.Windows;
using PropertyManager.App.BuildingLink;

namespace PropertyManager.App;
public partial class App : Application
{
    protected override async void OnStartup(StartupEventArgs e)
    {
        if(e.Args.Contains("--operations-scheduled"))
        {
            ShutdownMode=ShutdownMode.OnExplicitShutdown;base.OnStartup(e);int exit=1;
            try{await new Operations.OperationsRunner().RunAsync("preview",Operations.OperationsSettings.Load(),null,_=>{},true);exit=0;}catch(Exception){}
            Shutdown(exit);return;
        }
        if(e.Args.Contains("--buildinglink-scheduled"))
        {
            ShutdownMode=ShutdownMode.OnExplicitShutdown;base.OnStartup(e);
            int exit=1;
            try { exit=await new BuildingLinkRunner().RunAsync(_=>{}); } catch(Exception) { }
            Shutdown(exit);return;
        }
        base.OnStartup(e);new MainWindow().Show();
    }
}
