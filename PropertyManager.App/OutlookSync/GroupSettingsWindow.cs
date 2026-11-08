using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
namespace PropertyManager.App.OutlookSync;

internal sealed class GroupSettingsWindow : Window
{
    internal readonly Dictionary<string,Dictionary<string,TextBox>> Inputs=new();
    private readonly GroupNaming settings;
    internal GroupSettingsWindow(GroupNaming settings)
    {
        this.settings=settings;
        Title="Outlook group settings";Width=800;Height=720;MinWidth=580;MinHeight=430;WindowStartupLocation=WindowStartupLocation.CenterOwner;Background=new SolidColorBrush(Color.FromRgb(245,248,250));
        var root=new DockPanel{Margin=new Thickness(24)};Content=root;
        var heading=new TextBlock{Text="Group names for both corporations",FontSize=23,FontWeight=FontWeights.SemiBold,Margin=new Thickness(0,0,0,10)};DockPanel.SetDock(heading,Dock.Top);root.Children.Add(heading);
        var help=new TextBlock{Text="Change Outlook group names only; individual contact names are unaffected. Configure both corporations here, then save together. Keep the numbered placeholders. Use Update all corporations to apply saved settings; first-time group changes still need approval.",TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,16)};DockPanel.SetDock(help,Dock.Top);root.Children.Add(help);
        var actions=new WrapPanel{Margin=new Thickness(0,16,0,0)};DockPanel.SetDock(actions,Dock.Bottom);root.Children.Add(actions);
        var save=new Button{Content="Save group names",Margin=new Thickness(0,0,8,0)};save.SetResourceReference(BackgroundProperty,"ThemeAccent");save.SetResourceReference(BorderBrushProperty,"ThemeAccent");save.SetResourceReference(ForegroundProperty,"ThemeAccentForeground");
        var reset=new Button{Content="Restore all defaults",Margin=new Thickness(0,0,8,0)};var cancel=new Button{Content="Cancel",IsCancel=true};actions.Children.Add(save);actions.Children.Add(reset);actions.Children.Add(cancel);
        var tabs=new TabControl();tabs.SetResourceReference(StyleProperty,"FeatureTabs");root.Children.Add(tabs);
        foreach(string corporation in new[]{"DEMO 1001","DEMO 2002"})
        {
            var list=new StackPanel{Margin=new Thickness(0,6,0,0)};var inputs=new Dictionary<string,TextBox>();Inputs[corporation]=inputs;
            foreach(var (key,value) in settings.For(corporation).Where(p=>corporation=="DEMO 2002"?p.Key is "Residents" or "Owners" or "Block":p.Key!="Block"))
            {
                list.Children.Add(new TextBlock{Text=key+" group"+(key is "Block" or "Riser" or "Floor"?" naming pattern":""),FontWeight=FontWeights.SemiBold,Margin=new Thickness(0,0,0,6)});
                var input=new TextBox{Text=value,Padding=new Thickness(9),Margin=new Thickness(0,0,0,5)};inputs[key]=input;list.Children.Add(input);
                list.Children.Add(new TextBlock{Text="Default: "+GroupNaming.Defaults[key],FontSize=11,Foreground=Brushes.SlateGray,TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,17)});
            }
            tabs.Items.Add(new TabItem{Header=corporation,Content=new ScrollViewer{Content=list,Foreground=new SolidColorBrush(Color.FromRgb(23,43,58)),VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Padding=new Thickness(0,0,16,0)}});
        }
        reset.Click+=(_,_)=>{foreach(var inputs in Inputs.Values)foreach(var (key,input) in inputs)input.Text=GroupNaming.Defaults[key];};
        save.Click+=(_,_)=>{try{SaveSettings();DialogResult=true;}catch(Exception ex){MessageBox.Show(this,ex.Message,"Group names",MessageBoxButton.OK,MessageBoxImage.Warning);}};
    }
    internal void SaveSettings()
    {
        var candidate=new GroupNaming{Corporations=new(settings.Corporations)};var resolved=new List<string>();
        foreach(var (corporation,inputs) in Inputs)
        {
            var proposed=settings.For(corporation);foreach(var (key,input) in inputs)proposed[key]=input.Text.Trim();GroupNaming.Validate(proposed);
            candidate.Corporations[corporation]=proposed;
            var sample=new List<string>{corporation+" Resident Email List (BCC only)",corporation+" - Owners (BCC only)"};
            if(corporation=="DEMO 2002")sample.AddRange(Enumerable.Range(1,9).Select(i=>corporation+" - Block "+i+"00 (BCC only)"));
            else{sample.Add(corporation+" - PH Floor (BCC only)");sample.Add(corporation+" - All Retail Units (BCC only)");sample.AddRange(Enumerable.Range(1,99).Select(i=>corporation+" - Riser "+i.ToString("D2")+" (BCC only)"));sample.AddRange(Enumerable.Range(1,99).Select(i=>corporation+" - "+i+"th Floor (BCC only)"));}
            resolved.AddRange(candidate.Resolve(corporation,sample).Values);
        }
        if(resolved.Distinct(StringComparer.OrdinalIgnoreCase).Count()!=resolved.Count)throw new InvalidDataException("Two corporations would use the same Outlook group name. Give their groups distinct names before saving.");
        candidate.Save();
    }
}
