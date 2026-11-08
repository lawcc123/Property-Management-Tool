using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using PropertyManager.App.EmailLists;

namespace PropertyManager.App.BuildingLink;

internal sealed record CorporationConflict(string Corporation,FolderShortcut Download,FolderShortcut Lists);
internal sealed record CorporationDraft(BuildingLinkSettings Settings,List<CorporationConflict> Conflicts,List<string> Unsupported)
{
    internal List<FolderShortcut> Folders=>Settings.Buildings.Where(b=>b.OccupantFolder.Length>0||b.UnitFolder.Length>0).Select(b=>new FolderShortcut(b.Corporation,b.OccupantFolder,b.Color,b.UnitFolder)).ToList();
}
internal static class CorporationConfiguration
{
    internal static readonly string[] Names=["DEMO 1001","DEMO 2002"];
    internal static string LegacyPath=>Path.Combine(Services.AppSettings.DataDirectory,"occupant-folder-shortcuts.json");
    internal static string? Canonical(string name)=>Names.FirstOrDefault(c=>System.Text.RegularExpressions.Regex.IsMatch(name,@"\b"+c.Replace(" ",@"[\s_#-]*")+@"\b",System.Text.RegularExpressions.RegexOptions.IgnoreCase));
    internal static BuildingDestination Identity(string name)=>name switch
    {
        "DEMO 1001"=>new(){Corporation=name,PropertyName="Example Condominium - DEMO 1001",FacilityId="901001"},
        "DEMO 2002"=>new(){Corporation=name,PropertyName="Example Townhomes - DEMO 2002",FacilityId="902002"},
        _=>throw new InvalidDataException("Only DEMO 1001 and DEMO 2002 are supported.")
    };
    private static bool Same(string a,string b)=>a.TrimEnd('\\','/').Equals(b.TrimEnd('\\','/'),StringComparison.OrdinalIgnoreCase);
    internal static CorporationDraft Read(bool requireResolved=true)
    {
        var settings=Settings.SettingsDatabase.Read(BuildingLinkSettings.SettingsPath) is {} json?JsonSerializer.Deserialize<BuildingLinkSettings>(json)??throw new InvalidDataException("BuildingLink settings are empty."):new BuildingLinkSettings();
        if(settings.CorporationFoldersVersion>=1)return new(settings,[],[]);
        var legacy=File.Exists(LegacyPath)?JsonSerializer.Deserialize<List<FolderShortcut>>(File.ReadAllText(LegacyPath))??throw new InvalidDataException("Saved corporation folders are empty."):[];
        var conflicts=new List<CorporationConflict>();var unsupported=legacy.Select(f=>f.Name).Concat(settings.Buildings.Select(b=>b.Corporation)).Where(n=>Canonical(n) is null).Distinct().ToList();
        var merged=new List<BuildingDestination>();
        foreach(string name in Names)
        {
            var old=settings.Buildings.Where(b=>Canonical(b.Corporation)==name).ToList();var list=legacy.Where(f=>Canonical(f.Name)==name).ToList();
            if(old.Count>1||list.Count>1)throw new InvalidDataException("Duplicate saved entries for "+name+". Resolve the duplicate entries before migrating corporation folders.");
            var b=old.FirstOrDefault()??Identity(name);if(old.Count==0)b.Enabled=false;var f=list.FirstOrDefault();var identity=Identity(name);b.Corporation=name;b.PropertyName=identity.PropertyName;b.FacilityId=identity.FacilityId;
            b.Color=FolderColors.Resolve(f?.Color??b.Color,name);
            if(f is not null)
            {
                bool conflict=b.OccupantFolder.Length>0&&f.Path.Length>0&&!Same(b.OccupantFolder,f.Path)||b.UnitFolder.Length>0&&!string.IsNullOrWhiteSpace(f.UnitPath)&&!Same(b.UnitFolder,f.UnitPath);
                if(conflict)conflicts.Add(new(name,new(name,b.OccupantFolder,b.Color,b.UnitFolder),f));
                else{if(b.OccupantFolder.Length==0)b.OccupantFolder=f.Path;if(b.UnitFolder.Length==0)b.UnitFolder=f.UnitPath??"";}
            }
            merged.Add(b);
        }
        settings.Buildings=merged;
        if(requireResolved&&(conflicts.Count>0||unsupported.Count>0))throw new InvalidDataException("Existing corporation folder settings need a one-time review. Open BuildingLink Data > Manage folders to choose the shared paths.");
        return new(settings,conflicts,unsupported);
    }
    internal static void Save(BuildingLinkSettings settings)
    {
        var seen=new HashSet<string>();var paths=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach(var b in settings.Buildings)
        {
            if(!Names.Contains(b.Corporation)||!seen.Add(b.Corporation))throw new InvalidDataException("Choose each supported corporation only once.");
            var identity=Identity(b.Corporation);b.PropertyName=identity.PropertyName;b.FacilityId=identity.FacilityId;
            if(b.OccupantFolder.Length==0&&b.UnitFolder.Length==0){b.Enabled=false;continue;}
            var f=FolderShortcutStore.ValidateEntry(b.Corporation,b.OccupantFolder,[],b.Color,b.UnitFolder);b.OccupantFolder=f.Path;b.UnitFolder=f.UnitPath!;b.Color=f.Color;
            foreach(string path in new[]{b.OccupantFolder,b.UnitFolder})if(!paths.Add(Path.GetFullPath(path).TrimEnd('\\','/')))throw new InvalidDataException("Each corporation and report needs a separate folder.");
        }
        Directory.CreateDirectory(Services.AppSettings.DataDirectory);
        // Keep the original files once, before the first canonical save.
        if(!File.Exists(BuildingLinkSettings.SettingsPath)||JsonSerializer.Deserialize<BuildingLinkSettings>(Settings.SettingsDatabase.Read(BuildingLinkSettings.SettingsPath)??"{}")?.CorporationFoldersVersion!=1)
            foreach(string file in new[]{BuildingLinkSettings.SettingsPath,LegacyPath})if(File.Exists(file)&&!File.Exists(file+".before-shared-folders.bak"))File.Copy(file,file+".before-shared-folders.bak");
        settings.CorporationFoldersVersion=1;
        BuildingLinkSettings.AtomicWrite(BuildingLinkSettings.SettingsPath,JsonSerializer.Serialize(settings,new JsonSerializerOptions{WriteIndented=true}));
    }
    internal static void SaveFolders(IEnumerable<FolderShortcut> folders,BuildingLinkSettings? basis=null)
    {
        var settings=basis??Read().Settings;var entries=folders.ToList();var buildings=new List<BuildingDestination>();
        foreach(var f in entries)
        {
            var b=Identity(f.Name);var old=settings.Buildings.FirstOrDefault(v=>v.Corporation==f.Name);b.Enabled=old?.Enabled??true;b.OccupantFolder=f.Path;b.UnitFolder=f.UnitPath??"";b.Color=f.Color;buildings.Add(b);
        }
        settings.Buildings=buildings;Save(settings);
    }
    internal static bool Manage(Window owner,Func<bool>? beforeSave=null)
    {
        try
        {
            if(owner is MainWindow main&&main.CorporationSettingsBusy)throw new InvalidOperationException("Finish the current operation before changing shared corporation folders.");
            Directory.CreateDirectory(Services.AppSettings.DataDirectory);
            using var downloadLock=new FileStream(Path.Combine(Services.AppSettings.DataDirectory,"buildinglink-run.lock"),FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);
            var draft=Read(false);
            foreach(var conflict in draft.Conflicts)
            {
                var dialog=new Window{Owner=owner,Title="Choose shared corporation folders",Width=760,SizeToContent=SizeToContent.Height,WindowStartupLocation=WindowStartupLocation.CenterOwner,Background=System.Windows.Media.Brushes.White};
                var panel=new StackPanel{Margin=new Thickness(24)};dialog.Content=panel;
                panel.Children.Add(new TextBlock{Text=conflict.Corporation+" has different saved folders",FontSize=22,FontWeight=FontWeights.SemiBold,TextWrapping=TextWrapping.Wrap});
                panel.Children.Add(new TextBlock{Text="Choose the paths all features should use. Existing Excel files will stay in place.",Margin=new Thickness(0,12,0,16),TextWrapping=TextWrapping.Wrap});
                RadioButton Choice(string title,FolderShortcut f){var radio=new RadioButton{Content=title+"\nOccupant Data: "+f.Path+"\nUnit Data: "+f.UnitPath,Margin=new Thickness(0,0,0,20)};panel.Children.Add(radio);return radio;}
                var download=Choice("BuildingLink download destinations",conflict.Download);var lists=Choice("Email / owner / Outlook input folders",conflict.Lists);
                var buttons=new StackPanel{Orientation=Orientation.Horizontal,HorizontalAlignment=HorizontalAlignment.Right};panel.Children.Add(buttons);var cancel=new Button{Content="Cancel",IsCancel=true,Margin=new Thickness(0,0,10,0)};buttons.Children.Add(cancel);var next=new Button{Content="Use selected folders",IsEnabled=false};buttons.Children.Add(next);
                download.Checked+=(_,_)=>next.IsEnabled=true;lists.Checked+=(_,_)=>next.IsEnabled=true;next.Click+=(_,_)=>dialog.DialogResult=true;
                if(dialog.ShowDialog()!=true)return false;var chosen=download.IsChecked==true?conflict.Download:conflict.Lists;var b=draft.Settings.Buildings.Single(b=>b.Corporation==conflict.Corporation);b.OccupantFolder=chosen.Path;b.UnitFolder=chosen.UnitPath??"";
            }
            if(draft.Unsupported.Count>0&&MessageBox.Show(owner,"Only DEMO 1001 and DEMO 2002 are supported. These older entries will not be included in the shared configuration:\n\n"+string.Join("\n",draft.Unsupported)+"\n\nThe original settings will be retained in a recovery copy when you save. Continue?","Review existing corporations",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return false;
            var manager=new ManageFoldersWindow(draft.Folders){Owner=owner};if(manager.ShowDialog()!=true||beforeSave?.Invoke()==false)return false;
            SaveFolders(manager.SavedFolders,draft.Settings);return true;
        }
        catch(Exception ex){MessageBox.Show(owner,ex.Message,"Manage Condo Corporations",MessageBoxButton.OK,MessageBoxImage.Warning);return false;}
    }
}
