using System.IO;
using System.Text.Json;

namespace PropertyManager.App.EmailLists;

internal sealed record FolderShortcut(string Name, string Path, string? Color = null, string? UnitPath = null)
{
    [System.Text.Json.Serialization.JsonIgnore]
    public string ButtonColor => FolderColors.Resolve(Color, Name);
}

internal sealed record FolderColorOption(string Name, string Hex);
internal static class FolderColors
{
    internal static readonly FolderColorOption[] Options =
    [
        new("Green", "#147D70"), new("Pink", "#B43F69"), new("Red", "#B7443C"),
        new("Blue", "#3569A8"), new("Purple", "#7756A3"), new("Amber", "#94651A"), new("Slate", "#526575")
    ];
    internal static string Resolve(string? color, string name)
    {
        if (color is not null && Options.Any(option => option.Hex.Equals(color, StringComparison.OrdinalIgnoreCase)))
            return color.ToUpperInvariant();
        return System.Text.RegularExpressions.Regex.IsMatch(name, @"\bDEMO[\s_-]*1001\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            ? "#B43F69" : "#147D70";
    }
}

internal static class FolderShortcutStore
{
    public static FolderShortcut ValidateEntry(string name, string path, IEnumerable<FolderShortcut> otherFolders, string? color = null, string? unitPath = null)
    {
        name = name.Trim(); path = path.Trim();
        if (name.Length == 0 || name.Length > 60) throw new InvalidDataException("Enter a folder name between 1 and 60 characters.");
        if (!Path.IsPathFullyQualified(path)) throw new InvalidDataException("Choose a full local or network folder path.");
        path = Path.GetFullPath(path);
        if (otherFolders.Any(f => f.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException("That name is already used. Choose a different corporation or folder name.");
        if (otherFolders.Any(f => new[]{f.Path,f.UnitPath??""}.Any(p=>p.TrimEnd('\\', '/').Equals(path.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase))))
            throw new InvalidDataException("That backup folder is already saved. Select its existing button instead.");
        unitPath=(unitPath??"").Trim();
        if(!Path.IsPathFullyQualified(unitPath))throw new InvalidDataException("Choose a full local or network Unit Data folder path.");
        unitPath=Path.GetFullPath(unitPath);
        if(path.TrimEnd('\\','/').Equals(unitPath.TrimEnd('\\','/'),StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("Choose separate Occupant Data and Unit Data folders.");
        if(otherFolders.Any(f=>new[]{f.Path,f.UnitPath??""}.Any(p=>p.TrimEnd('\\','/').Equals(unitPath.TrimEnd('\\','/'),StringComparison.OrdinalIgnoreCase))))throw new InvalidDataException("That Unit Data folder is already used by another corporation.");
        return new(name, path, FolderColors.Resolve(color, name),unitPath);
    }
    private static string FilePath => Path.Combine(PropertyManager.App.Services.AppSettings.DataDirectory, "occupant-folder-shortcuts.json");
    public static List<FolderShortcut> Load(string? filePath = null)
    {
        if(filePath is null)return BuildingLink.CorporationConfiguration.Read().Folders;
        string path = filePath;
        if (!File.Exists(path)) return [];
        var folders=(JsonSerializer.Deserialize<List<FolderShortcut>>(File.ReadAllText(path)) ?? [])
            .Where(s => s is not null && !string.IsNullOrWhiteSpace(s.Name) && !string.IsNullOrWhiteSpace(s.Path)).ToList();
        return folders;
    }

    public static void Save(IEnumerable<FolderShortcut> shortcuts, string? filePath = null)
    {
        if(filePath is null){BuildingLink.CorporationConfiguration.SaveFolders(shortcuts);return;}
        string path = filePath;
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            File.WriteAllText(temp, JsonSerializer.Serialize(shortcuts, new JsonSerializerOptions { WriteIndented = true }));
            File.Move(temp, path, overwrite: true);
        }
        finally { if (File.Exists(temp)) File.Delete(temp); }
    }
}
