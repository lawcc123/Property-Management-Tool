using System.Data;
using System.IO;
using PropertyManager.App.EmailLists;
namespace PropertyManager.App.PersonalizedEmail;
internal sealed record SnapshotImport(RecipientSnapshotSession Session,int ChangedCells,int ChangedEmails);
internal sealed class RecipientSnapshotSession
{
    internal const string Id="Recipient ID";
    internal static readonly string[] Core=[Id,"Unit","Names","Include","Corporation","Email"];
    internal DataTable Table {get;}=new(){CaseSensitive=false};
    private readonly Dictionary<string,CampaignRecipient> originals;
    internal RecipientSnapshotSession(IEnumerable<CampaignRecipient> recipients)
    {
        var rows=recipients.ToList();
        if(rows.Any(r=>string.IsNullOrWhiteSpace(r.RecipientId))||rows.GroupBy(r=>r.RecipientId,StringComparer.OrdinalIgnoreCase).Any(g=>g.Count()>1))throw new InvalidDataException("Recipient IDs are missing or duplicated. Load the recipient snapshot again.");
        originals=rows.ToDictionary(r=>r.RecipientId,StringComparer.OrdinalIgnoreCase);
        foreach(string key in Core)Table.Columns.Add(key,key=="Include"?typeof(bool):typeof(string));
        foreach(string key in rows.SelectMany(r=>r.Fields.Keys).Distinct(StringComparer.OrdinalIgnoreCase).Where(k=>!Core.Contains(k,StringComparer.OrdinalIgnoreCase)&&!k.Equals("Status",StringComparison.OrdinalIgnoreCase)))AddColumn(key);
        foreach(var recipient in rows)
        {
            var row=Table.NewRow();row[Id]=recipient.RecipientId;row["Unit"]=recipient.Unit;row["Names"]=recipient.Names;row["Include"]=recipient.Included;row["Corporation"]=recipient.Corporation;row["Email"]=recipient.Email;
            foreach(DataColumn column in Table.Columns)if(!Core.Contains(column.ColumnName))row[column]=recipient.Fields.GetValueOrDefault(column.ColumnName,"");
            Table.Rows.Add(row);
        }
        Table.Columns[Id]!.ReadOnly=true;
    }
    internal static string ValidateColumn(string value)
    {
        string name=value.Trim();
        if(name.Length==0||name.Length>80||name.Any(c=>char.IsControl(c)||c is '{' or '}'))throw new InvalidDataException("Use a column name of 1 to 80 characters without braces or line breaks.");
        if(Core.Contains(name,StringComparer.OrdinalIgnoreCase)||name.Equals("Status",StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("This name is reserved for an existing recipient field.");
        return name;
    }
    internal void AddColumn(string name)
    {
        name=ValidateColumn(name);
        if(Table.Columns.Cast<DataColumn>().Any(c=>c.ColumnName.Equals(name,StringComparison.OrdinalIgnoreCase)))throw new InvalidDataException("A column with this name already exists.");
        var column=Table.Columns.Add(name,typeof(string));column.DefaultValue="";
        foreach(DataRow row in Table.Rows)row[column]="";
    }
    internal List<CampaignRecipient> Recipients()=>Table.Rows.Cast<DataRow>().Select(row=>
    {
        var original=originals[(string)row[Id]];
        var fields=new Dictionary<string,string>(original.Fields,StringComparer.OrdinalIgnoreCase);
        foreach(DataColumn c in Table.Columns)if(!Core.Contains(c.ColumnName))fields[c.ColumnName]=Convert.ToString(row[c])??"";
        return new CampaignRecipient{RecipientId=original.RecipientId,Unit=(string)row["Unit"],Names=(string)row["Names"],Corporation=(string)row["Corporation"],Email=(string)row["Email"],Included=(bool)row["Include"],Fields=fields};
    }).ToList();
    internal RecipientFileData Export()
    {
        var columns=Table.Columns.Cast<DataColumn>().Select(c=>c.ColumnName).ToList();
        return new(columns,Table.Rows.Cast<DataRow>().Select(row=>columns.ToDictionary(c=>c,c=>c=="Include"?((bool)row[c]?"Yes":"No"):Convert.ToString(row[c])??"",StringComparer.OrdinalIgnoreCase)).ToList());
    }
    internal SnapshotImport Import(RecipientFileData data)
    {
        if(data.Columns.Distinct(StringComparer.OrdinalIgnoreCase).Count()!=data.Columns.Count||data.Columns.Any(c=>c!=c.Trim()))throw new InvalidDataException("Keep unique column headings without leading or trailing spaces.");
        if(Core.Any(c=>!data.Columns.Contains(c,StringComparer.OrdinalIgnoreCase)))throw new InvalidDataException("Import the exported snapshot with Recipient ID, Unit, Names, Include, Corporation and Email columns intact.");
        string Value(Dictionary<string,string> row,string key)=>row.FirstOrDefault(p=>p.Key.Equals(key,StringComparison.OrdinalIgnoreCase)).Value??"";
        var ids=data.Rows.Select(r=>Value(r,Id)).ToList();
        if(ids.Any(string.IsNullOrWhiteSpace)||ids.Distinct(StringComparer.OrdinalIgnoreCase).Count()!=ids.Count)throw new InvalidDataException("Recipient IDs are missing or duplicated. No changes were applied.");
        if(ids.Count!=originals.Count||ids.Any(id=>!originals.ContainsKey(id)))throw new InvalidDataException("The recipient IDs do not match this snapshot. Keep every exported row and use Include = No to exclude recipients.");
        var next=new RecipientSnapshotSession(Recipients());
        foreach(string name in data.Columns.Where(c=>!Core.Contains(c,StringComparer.OrdinalIgnoreCase)))
        {
            ValidateColumn(name);
            if(!next.Table.Columns.Contains(name))next.AddColumn(name);
        }
        int changes=0,emails=0;
        foreach(var input in data.Rows)
        {
            var target=next.Table.Rows.Cast<DataRow>().Single(r=>string.Equals((string)r[Id],Value(input,Id),StringComparison.OrdinalIgnoreCase));
            foreach(string name in data.Columns.Where(c=>!c.Equals(Id,StringComparison.OrdinalIgnoreCase)))
            {
                object value=Value(input,name);
                if(name.Equals("Include",StringComparison.OrdinalIgnoreCase))value=((string)value).Trim().ToLowerInvariant() switch{"yes" or "true" or "y" or "1"=>true,"no" or "false" or "n" or "0"=>false,_=>throw new InvalidDataException("Include must be Yes or No for unit "+Value(input,"Unit"))};
                if(!Equals(target[name],value)){changes++;if(name.Equals("Email",StringComparison.OrdinalIgnoreCase))emails++;target[name]=value;}
            }
        }
        return new(next,changes,emails);
    }
}
