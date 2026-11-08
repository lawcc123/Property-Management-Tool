using System.IO;
using System.Text.Json;

namespace PropertyManager.App.BuildingLink;

internal sealed class BuildingDestination : System.ComponentModel.INotifyPropertyChanged
{
    public string Corporation { get; set; } = "";
    public string PropertyName { get; set; } = "";
    public string FacilityId { get; set; } = "";
    private string occupantFolder="",unitFolder="";
    public string OccupantFolder {get=>occupantFolder;set{occupantFolder=value;PropertyChanged?.Invoke(this,new(nameof(OccupantFolder)));}}
    public string UnitFolder {get=>unitFolder;set{unitFolder=value;PropertyChanged?.Invoke(this,new(nameof(UnitFolder)));}}
    public string? Color {get;set;}
    [System.Text.Json.Serialization.JsonIgnore]
    public string ButtonColor=>EmailLists.FolderColors.Resolve(Color,Corporation);
    public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
    public bool Enabled { get; set; } = false;
}

internal sealed class BuildingLinkSettings
{
    public int CorporationFoldersVersion {get;set;}
    public List<BuildingDestination> Buildings { get; set; } =
    [
        new() { Corporation="DEMO 2002", PropertyName="Example Townhomes - DEMO 2002", FacilityId="902002" },
        new() { Corporation="DEMO 1001", PropertyName="Example Condominium - DEMO 1001", FacilityId="901001" }
    ];
    public string Time { get; set; } = "06:00";
    public List<int> Days { get; set; } = [0,1,2,3,4,5,6];
    internal static string SettingsPath => Path.Combine(Services.AppSettings.DataDirectory,"buildinglink-settings.json");
    internal static string ProgressPath => Path.Combine(Services.AppSettings.DataDirectory,"buildinglink-progress.json");
    internal static string SuccessPath => Path.Combine(Services.AppSettings.DataDirectory,"buildinglink-last-success.json");
    internal static string ResultPath => Path.Combine(Services.AppSettings.DataDirectory,"buildinglink-last-run.json");
    internal static BuildingLinkSettings Load()=>CorporationConfiguration.Read().Settings;
    internal void Validate()
    {
        if (!Buildings.Any(b=>b.Enabled)) throw new InvalidDataException("Select at least one corporation.");
        foreach(var b in Buildings.Where(b=>b.Enabled))
        {
            if (string.IsNullOrWhiteSpace(b.Corporation) || string.IsNullOrWhiteSpace(b.PropertyName) || !int.TryParse(b.FacilityId,out _))
                throw new InvalidDataException("The corporation's BuildingLink identity is incomplete.");
            foreach(var path in new[]{b.OccupantFolder,b.UnitFolder})
                if(string.IsNullOrWhiteSpace(path) || !Path.IsPathFullyQualified(path))
                    throw new InvalidDataException($"Choose full Occupant Data and Unit Data folder paths for {b.Corporation}.");
            if (Path.GetFullPath(b.OccupantFolder).TrimEnd('\\').Equals(Path.GetFullPath(b.UnitFolder).TrimEnd('\\'),StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException($"Choose separate Occupant Data and Unit Data folders for {b.Corporation}.");
        }
        var destinations=Buildings.Where(b=>b.Enabled).SelectMany(b=>new[]{b.OccupantFolder,b.UnitFolder}).Select(p=>Path.GetFullPath(p).TrimEnd('\\')).ToList();
        if(destinations.Distinct(StringComparer.OrdinalIgnoreCase).Count()!=destinations.Count)
            throw new InvalidDataException("Each corporation and report needs a separate destination folder.");
    }
    internal void Save()=>CorporationConfiguration.Save(this);
    internal static void AtomicWrite(string path,string contents)
    {
        Settings.SettingsDatabase.Write(path,contents);
    }
}
