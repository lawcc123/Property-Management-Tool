using System.IO;
using System.Text.RegularExpressions;

namespace PropertyManager.App.EmailLists;

internal sealed record AlternateMailingAddress
{
    public string Name {get;init;}="";
    public string Line1 {get;init;}="";
    public string Line2 {get;init;}="";
    public string City {get;init;}="";
    public string Province {get;init;}="";
    public string Postal {get;init;}="";
    public string Country {get;init;}="";
    public string Status {get;init;}="";
    public string Text=>string.Join("\n",new[]{Line1,Line2,string.Join(" ",new[]{string.Join(", ",new[]{City,Province}.Where(v=>v.Length>0)),Postal}.Where(v=>v.Length>0)),Country}.Where(v=>v.Length>0));
    public string Key=>string.Join("\u001F",new[]{Line1,Line2,City,Province,Postal,Country}.Select(v=>v.ToUpperInvariant()));
    public string[] ExportValues=>[Name,Text,Status];
    internal static readonly string[] ExportColumns=["Mailing Name","Mailing Address","Mailing Address Status"];
    private static readonly string[][] Aliases=
    [ ["Alt Mail Name","Alternate Mailing Name","Alternate Name"],
      ["Alt Mail Address Line 1","Alternate Address Line 1","Alternate Address","Alternate Mailing Address","Alternate Mailing Address Line 1","Alt Mailing Address Line 1","Alternate Address 1"],
      ["Alt Mail Address Line 2","Alternate Address Line 2","Alternate Mailing Address Line 2","Alt Mailing Address Line 2","Alternate Address 2"],
      ["Alt Mail City","Alternate City","Alternate Mailing City"],
      ["Alt Mail State","Alternate State","Alternate Province","Alternate Mailing State","Alternate Mailing Province"],
      ["Alt Mail Zip","Alternate Zip","Alternate Postal Code","Alternate Mailing Postal Code","Alternate Mailing Zip"],
      ["Alt Mail Country Code","Alternate Country","Alternate Country Code","Alternate Mailing Country"] ];
    private static string Normalize(string text)=>Regex.Replace(text,@"[^\p{L}\p{Nd}]","").ToUpperInvariant();
    internal static string UnitKey(string unit)=>Regex.Replace(unit.Trim(),@"\s+","").ToUpperInvariant();
    private static string Field(Dictionary<string,string> row,IEnumerable<string> names)
    {
        var keys=names.Select(Normalize).ToHashSet();
        return row.Where(pair=>keys.Contains(Normalize(pair.Key))).Select(p=>p.Value.Trim()).FirstOrDefault(v=>v.Length>0)??"";
    }
    internal static bool HasColumns(RecipientFileData data)=>data.Columns.Any(c=>Aliases[1].Any(a=>Normalize(a)==Normalize(c)));
    internal static Dictionary<string,Dictionary<string,string>> UnitMap(RecipientFileData data)
    {
        if(!HasColumns(data)||!data.Columns.Any(c=>new[]{"UNITID","UNITIDNUMBER"}.Contains(Normalize(c))))
            throw new InvalidDataException("Choose Unit Data containing UnitID # (or Unit ID) and alternate mailing-address columns.");
        var map=new Dictionary<string,Dictionary<string,string>>(StringComparer.OrdinalIgnoreCase);
        foreach(var row in data.Rows)
        {
            string key=UnitKey(Field(row,["UnitID #","Unit ID","Unit ID Number"]));if(key.Length==0)continue;
            if(!map.TryAdd(key,row))throw new InvalidDataException("Unit Data contains duplicate unit identifiers. Choose a single corporation's export.");
        }
        if(map.Count==0)throw new InvalidDataException("Unit Data contains no unit records.");
        return map;
    }
    internal static AlternateMailingAddress Read(Dictionary<string,string> owner,Dictionary<string,Dictionary<string,string>>? units)
    {
        var row=owner;string source="occupant data";
        bool any=Aliases.Skip(1).Any(names=>Field(owner,names).Length>0);
        if(!any&&units is not null)
        {
            if(!units.TryGetValue(UnitKey(GeneratedEmailListService.Value(owner,"Unit ID")),out var unit))return new(){Status="Unit not found in Unit Data - review"};
            row=unit;source="Unit Data";
        }
        var values=Aliases.Select(names=>Field(row,names)).ToArray();
        bool offsite=new[]{"Offsite Owner","Off-site Owner"}.Contains(GeneratedEmailListService.Value(owner,"User Type"),StringComparer.OrdinalIgnoreCase);
        var missing=new[]{(1,"street"),(3,"city"),(4,"province/state"),(5,"postal/ZIP")}.Where(p=>values[p.Item1].Length==0).Select(p=>p.Item2).ToArray();
        string status=values.Skip(1).All(v=>v.Length==0)?(offsite?"Missing alternate address - review":"No alternate address supplied")
            :missing.Length>0?"Review: missing "+string.Join(", ",missing):"Available ("+source+")";
        return new(){Name=values[0],Line1=values[1],Line2=values[2],City=values[3],Province=values[4],Postal=values[5],Country=values[6],Status=status};
    }
}
