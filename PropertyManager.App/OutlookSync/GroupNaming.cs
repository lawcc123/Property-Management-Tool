using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace PropertyManager.App.OutlookSync;

internal sealed class GroupNaming
{
    public Dictionary<string,Dictionary<string,string>> Corporations {get;set;}=new();
    internal static readonly Dictionary<string,string> Defaults=new(){
        ["Residents"]="{corporation} Resident Email List (BCC only)",
        ["Owners"]="{corporation} - Owners (BCC only)",
        ["Block"]="{corporation} - Block {block} (BCC only)",
        ["Riser"]="{corporation} - Riser {riser} (BCC only)",
        ["Floor"]="{corporation} - {floor} Floor (BCC only)",
        ["Penthouse"]="{corporation} - PH Floor (BCC only)",
        ["Retail"]="{corporation} - All Retail Units (BCC only)"};
    internal static string SettingsPath=>Path.Combine(Services.AppSettings.DataDirectory,"OutlookSync","group-names.json");
    internal static GroupNaming Load(string? path=null)=>Settings.SettingsDatabase.Read(path??SettingsPath) is {} json?JsonSerializer.Deserialize<GroupNaming>(json)??throw new InvalidDataException("Group name settings are empty."):new();
    internal Dictionary<string,string> For(string corp)=>Defaults.ToDictionary(p=>p.Key,p=>Corporations.GetValueOrDefault(corp)?.GetValueOrDefault(p.Key)??p.Value);
    internal static void Validate(Dictionary<string,string> names)
    {
        foreach(var (key,defaultValue) in Defaults)
        {
            if(!names.TryGetValue(key,out string? value)||string.IsNullOrWhiteSpace(value)||value.Length>200||value.Any(char.IsControl))throw new InvalidDataException("Enter a group name of 1 to 200 characters for "+key+".");
            var allowed=Regex.Matches(defaultValue,@"\{[^}]+\}").Select(m=>m.Value).ToHashSet();
            if(Regex.Matches(value,@"\{[^}]+\}").Any(m=>!allowed.Contains(m.Value)))throw new InvalidDataException("Unknown placeholder in "+key+". Use the placeholders shown in its default name.");
            if(Regex.Replace(value,@"\{[^}]+\}","").IndexOfAny(['{','}'])>=0)throw new InvalidDataException("Unmatched placeholder braces in "+key+".");
            foreach(string token in allowed.Where(t=>t!="{corporation}"))if(!value.Contains(token,StringComparison.Ordinal))throw new InvalidDataException(key+" must include "+token+" so each group has a distinct name.");
        }
    }
    internal void Save(string? path=null)
    {
        foreach(var names in Corporations.Values)Validate(names);
        path??=SettingsPath;Settings.SettingsDatabase.Write(path,JsonSerializer.Serialize(this,new JsonSerializerOptions{WriteIndented=true}));
    }
    internal Dictionary<string,string> Resolve(string corp,IEnumerable<string> canonical)
    {
        var names=For(corp);Validate(names);var result=new Dictionary<string,string>();
        foreach(string key in canonical)
        {
            string suffix=key[corp.Length..];string kind,token="",value="";
            if(suffix==" Resident Email List (BCC only)")kind="Residents";
            else if(suffix==" - Owners (BCC only)")kind="Owners";
            else if(suffix==" - PH Floor (BCC only)")kind="Penthouse";
            else if(suffix==" - All Retail Units (BCC only)")kind="Retail";
            else if(suffix.StartsWith(" - Block ")){kind="Block";token="{block}";value=suffix[9..^11];}
            else if(suffix.StartsWith(" - Riser ")){kind="Riser";token="{riser}";value=suffix[9..^11];}
            else{kind="Floor";token="{floor}";value=suffix[3..^17];}
            string name=names[kind].Trim().Replace("{corporation}",corp,StringComparison.Ordinal);
            if(token.Length>0)name=name.Replace(token,value,StringComparison.Ordinal);
            result[key]=name;
        }
        if(result.Values.Distinct(StringComparer.OrdinalIgnoreCase).Count()!=result.Count)throw new InvalidDataException("Two Outlook groups would have the same name. Change Group settings before previewing.");
        return result;
    }
}
