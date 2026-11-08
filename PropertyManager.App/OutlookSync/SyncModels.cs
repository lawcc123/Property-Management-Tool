using System.Text.Json;

namespace PropertyManager.App.OutlookSync;

internal sealed record OutlookMailbox(string StoreId,string FolderId,string Label);
internal sealed record ResidentLink(string Corporation,List<string> Units,List<string> Names);
internal sealed record GroupMember(string Email,string Name);
internal sealed record OutlookEntity
{
    public List<string> ReadIssues {get;init;}=[];
    public string Id {get;init;}="";
    public byte[]? FullItemBackup {get;init;}
    public string Kind {get;init;}="Contact";
    public string Modified {get;init;}="";
    // Optional for compatibility with recovery snapshots written before capture
    // of Outlook's FileUnderId rule was introduced. null + captured means absent.
    public bool FilingFormatCaptured {get;init;}
    public int? FilingFormat {get;init;}
    public string ManagedId {get;init;}="";
    public string SourceGroup {get;init;}="";
    public string MembersFingerprint {get;init;}="";
    public bool Owned {get;init;}
    public List<ResidentLink> Links {get;init;}=[];
    public Dictionary<string,string> Fields {get;init;}=new(){{"FullName",""},{"FileAs",""},{"Email1DisplayName",""},{"Email2DisplayName",""},{"Email3DisplayName",""}};
    public List<string> Emails {get;init;}=["","",""];
    public List<GroupMember> Members {get;init;}=[];
    public string Name=>Fields.GetValueOrDefault("FullName","");
    public bool Managed=>ManagedId.Length>0;
    private string Content()=>JsonSerializer.Serialize(new{Kind,ManagedId,Owned,SourceGroup,MembersFingerprint,
        Links=Links.OrderBy(l=>l.Corporation).Select(l=>new{l.Corporation,Units=l.Units.Order(),Names=l.Names.Order()}),
        Fields=Fields.OrderBy(f=>f.Key),Emails,ReadIssues,Members=Members.OrderBy(m=>m.Email).ThenBy(m=>m.Name)});
    internal bool SameContent(OutlookEntity? other)=>other is not null&&Content()==other.Content()&&(!FilingFormatCaptured||!other.FilingFormatCaptured||FilingFormat==other.FilingFormat);
    internal bool SameVersion(OutlookEntity? other)=>SameContent(other)&&Id==other!.Id&&Modified==other.Modified;
    internal bool SameSyncState(OutlookEntity? other)=>other is not null&&Id==other.Id&&SameContent(other);
    internal List<string> ChangedFields(OutlookEntity? other)
    {
        if(other is null)return ["item missing or moved"];
        var differences=Fields.Keys.Union(other.Fields.Keys).Where(k=>Fields.GetValueOrDefault(k,"")!=other.Fields.GetValueOrDefault(k,"")).ToList();
        for(int i=0;i<Math.Max(Emails.Count,other.Emails.Count);i++)if(Emails.ElementAtOrDefault(i)!=other.Emails.ElementAtOrDefault(i))differences.Add("Email"+(i+1)+" address");
        if(JsonSerializer.Serialize(Links)!=JsonSerializer.Serialize(other.Links)||ManagedId!=other.ManagedId||Owned!=other.Owned||SourceGroup!=other.SourceGroup||MembersFingerprint!=other.MembersFingerprint)differences.Add("app associations");
        if(JsonSerializer.Serialize(Members.OrderBy(m=>m.Email).ThenBy(m=>m.Name))!=JsonSerializer.Serialize(other.Members.OrderBy(m=>m.Email).ThenBy(m=>m.Name)))differences.Add("group members");
        if(FilingFormatCaptured&&other.FilingFormatCaptured&&FilingFormat!=other.FilingFormat)differences.Add("filing-name format");
        if(!ReadIssues.SequenceEqual(other.ReadIssues))differences.Add("address resolution");
        if(differences.Count==0&&Modified!=other.Modified)differences.Add("modification timestamp only");
        return differences;
    }
}
internal sealed class SyncChange
{
    public string Corporation {get;set;}="";
    public bool RestoreOriginal {get;init;}
    public bool Selected {get;set;}
    public bool NeedsManualSelection {get;init;}
    public bool Recommended=>CanApply&&(RestoreOriginal||!NeedsManualSelection&&After is not null&&(Before is null||Before.Owned));
    public bool CanApply {get;init;}=true;
    public bool RequiresApproval=>CanApply&&Before is {Kind:"Contact",Owned:false}&&After is not null;
    public string Action {get;init;}="";
    public string Email {get;init;}="";
    public string Detail {get;init;}="";
    public string Reason=>Action=="Unchanged"?"":Detail;
    public OutlookEntity? Before {get;init;}
    public OutlookEntity? After {get;init;}
    public string ItemKind=>(After??Before)?.Kind??"Review";
    public string CurrentName=>Before?.Name??"";
    public string ProposedName=>!CanApply&&Action!="Unchanged"?"Review required":After is null?(Before is null?"Review required":"(remove)"):After.Kind=="Contact"?After.Fields.GetValueOrDefault("FileAs",After.Name):After.Name;
    public string CurrentFiling=>Before?.Fields.GetValueOrDefault("FileAs","")??"";
    public string DisplayField {get;init;}="Email1DisplayName";
    public List<string> ExtraDisplayFields {get;init;}=[];
    public IEnumerable<string> DisplayFields=>new[]{DisplayField}.Concat(ExtraDisplayFields);
    public string CurrentDisplay=>string.Join("; ",DisplayFields.Select(f=>f+": "+(Before?.Fields.GetValueOrDefault(f,"")??"")));
    public string ProposedDisplay=>string.Join("; ",DisplayFields.Select(f=>f+": "+(After?.Fields.GetValueOrDefault(f,"")??"")));
    internal OutlookEntity? ApprovedAfter()
    {
        if(!CanApply)throw new InvalidOperationException("Resolve this review item before applying it.");
        if(RestoreOriginal||!RequiresApproval)return After;
        var fields=new Dictionary<string,string>(Before!.Fields);
        if(Selected)
            foreach(var key in new[]{"FileAs"}.Concat(DisplayFields))
                fields[key]=After!.Fields.GetValueOrDefault(key,"");
        return After! with {Fields=fields,FilingFormatCaptured=Selected?After.FilingFormatCaptured:Before.FilingFormatCaptured,FilingFormat=Selected?After.FilingFormat:Before.FilingFormat};
    }
}
internal sealed record SyncPlan(OutlookMailbox Mailbox,string Corporation,string Mode,List<SyncChange> Changes,string? SourcePath=null,string? SourceHash=null);
internal sealed record SourceTargets(Dictionary<string,ResidentLink> Contacts,Dictionary<string,List<GroupMember>> Groups,List<string> Warnings)
{
    public Dictionary<string,string> GroupNames {get;init;}=new();
}
internal interface IOutlookGateway
{
    List<OutlookMailbox> Mailboxes();
    List<OutlookEntity> Read(OutlookMailbox mailbox,Action<int,int>? progress=null,CancellationToken cancellation=default);
    OutlookEntity CaptureBackup(OutlookMailbox mailbox,OutlookEntity before);
    OutlookEntity? Write(OutlookMailbox mailbox,OutlookEntity? before,OutlookEntity? after);
}

// Optional capability. Recovery/undo and gateways without it keep full scans.
internal interface IOutlookTargetedGateway
{
    List<OutlookEntity> ReadForChanges(OutlookMailbox mailbox,
        IReadOnlyList<(OutlookEntity? Before,OutlookEntity? After)> changes,
        Action<int,int>? progress=null,CancellationToken cancellation=default);
}
