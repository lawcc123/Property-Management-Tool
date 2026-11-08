using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using PropertyManager.App.EmailLists;

namespace PropertyManager.App.OutlookSync;

internal static class SyncPlanner
{
    internal static string Corporation(string value)
    {
        var m=Regex.Match(value,@"DEMO[\s_#-]*(2002|1001)\b",RegexOptions.IgnoreCase);
        if(!m.Success)throw new InvalidDataException("Choose DEMO 2002 or DEMO 1001 for this Outlook sync.");
        return "DEMO "+m.Groups[1].Value;
    }
    internal static SourceTargets Targets(RecipientFileData data,string corporation)
    {
        GeneratedEmailListService.Validate(data);corporation=Corporation(corporation);
        var buildings=data.Rows.Select(r=>GeneratedEmailListService.Value(r,"Building Name")).Where(v=>v.Length>0).Distinct().ToList();
        if(buildings.Count!=1||Corporation(buildings[0])!=corporation)throw new InvalidDataException("The workbook's corporation does not match the selection.");
        var people=new Dictionary<string,List<(string Unit,string Name)>>();
        var groups=new Dictionary<string,Dictionary<string,List<(string Unit,string Name)>>>();
        string main=corporation+" Resident Email List (BCC only)",owners=corporation+" - Owners (BCC only)";
        groups[main]=new();groups[owners]=new();var warnings=new List<string>();
        foreach(var row in data.Rows)
        {
            string V(string k)=>GeneratedEmailListService.Value(row,k);
            if(!V("User Active?").Equals("Y",StringComparison.OrdinalIgnoreCase)||!V("Unit Active?").Equals("Y",StringComparison.OrdinalIgnoreCase))continue;
            string unit=V("Unit ID"),name=string.Join(" ",new[]{V("First Name"),V("Last Name")}.Where(x=>x.Length>0));
            if(unit.Length==0)throw new InvalidDataException("An active occupant has no unit number. Fix the source data before syncing.");
            bool retail=unit.StartsWith("Retail",StringComparison.OrdinalIgnoreCase);
            var keys=new List<string>();if(!(corporation=="DEMO 1001"&&retail))keys.Add(main);
            string type=Regex.Replace(V("User Type"),@"[\s-]","").ToLowerInvariant();if(type is "owner" or "offsiteowner")keys.Add(owners);
            if(corporation=="DEMO 2002"&&Regex.IsMatch(unit,@"^\d+$"))keys.Add(corporation+" - Block "+unit[0]+"00 (BCC only)");
            if(corporation=="DEMO 1001")
            {
                if(retail)keys.Add(corporation+" - All Retail Units (BCC only)");
                else if(Regex.IsMatch(unit,@"^\d{3,4}$")){keys.Add(corporation+" - Riser "+unit[^2..]+" (BCC only)");keys.Add(corporation+" - "+Ordinal(int.Parse(unit[..^2]))+" Floor (BCC only)");}
                else {var ph=Regex.Match(unit,@"^PH\s*(\d{2})$",RegexOptions.IgnoreCase);if(ph.Success){keys.Add(corporation+" - Riser "+ph.Groups[1].Value+" (BCC only)");keys.Add(corporation+" - PH Floor (BCC only)");}}
            }
            foreach(string raw in GeneratedEmailListService.EmailParts(V("Emails")))
            {
                string email=raw.Trim().ToLowerInvariant();
                if(!GeneratedEmailListService.IsValidEmail(email)){warnings.Add("Unit "+unit+": invalid email skipped: "+raw);continue;}
                if(!people.ContainsKey(email))people[email]=[];people[email].Add((unit,name));
                foreach(string key in keys){if(!groups.ContainsKey(key))groups[key]=new();if(!groups[key].ContainsKey(email))groups[key][email]=[];groups[key][email].Add((unit,name));}
            }
        }
        if(people.Count==0)throw new InvalidDataException("No active occupants with valid emails. Outlook changes are blocked.");
        if(warnings.Count>0)throw new InvalidDataException("Invalid source emails must be corrected before syncing, to avoid removing valid Outlook entries. "+string.Join("; ",warnings.Take(5)));
        ResidentLink Link(List<(string Unit,string Name)> rows)=>new(corporation,rows.Select(r=>r.Unit).Where(v=>v.Length>0).Distinct().Order().ToList(),rows.Select(r=>r.Name).Where(v=>v.Length>0).Distinct().Order().ToList());
        return new(people.ToDictionary(k=>k.Key,k=>Link(k.Value)),groups.ToDictionary(k=>k.Key,k=>k.Value.OrderBy(x=>x.Key).Select(x=>new GroupMember(x.Key,Display([Link(x.Value)]))).ToList()),warnings);
    }
    private static string Ordinal(int n)=>n+(n%100 is 11 or 12 or 13?"th":(n%10) switch{1=>"st",2=>"nd",3=>"rd",_=>"th"});
    internal static string Display(IEnumerable<ResidentLink> links)=>string.Join("; ",links.OrderBy(l=>l.Corporation).Select(l=>string.Join(", ",l.Units.Select(u=>l.Corporation=="DEMO 2002"&&!u.StartsWith("TH",StringComparison.OrdinalIgnoreCase)?"TH "+u:u))+" - "+string.Join(" & ",l.Names)+" - "+l.Corporation));
    private static OutlookEntity New(string kind)=>new(){Kind=kind,ManagedId=Guid.NewGuid().ToString("N"),Owned=true};
    private static OutlookEntity? FilingPolicy(OutlookEntity? before,OutlookEntity? after)
    {
        // Only pin a custom rule when the approved filing text changes. Unchanged
        // contacts keep their current rule; old snapshots retain unknown metadata.
        if(after?.Kind=="Contact"&&(before is null||before.FilingFormatCaptured)&&
            (before is null||before.Fields.GetValueOrDefault("FileAs","")!=after.Fields.GetValueOrDefault("FileAs","")))
            return after with{FilingFormatCaptured=true,FilingFormat=-1};
        return after;
    }
    private static SyncChange Change(OutlookEntity? before,OutlookEntity? after,string email,string detail="")
    {
        after=FilingPolicy(before,after);
        bool same=before?.SameContent(after)==true;
        if(!same&&before is not null&&after is not null)detail+=" Changes: "+string.Join(", ",before.ChangedFields(after).Where(f=>f!="modification timestamp only"))+".";
        bool approval=before is {Owned:false};
        bool large=before is {Kind:"Group"}&&before.Members.Count>=5&&after is not null&&before.Members.Count(m=>!after.Members.Any(n=>n.Email==m.Email))>=before.Members.Count*0.5;
        if(large)detail+=" Large membership reduction - manual selection required.";
        return new(){Before=before,After=after,Email=email,Selected=!same&&!approval&&!large&&after is not null,NeedsManualSelection=large,CanApply=!same,Action=same?"Unchanged":before is null?"Create":after is null?"Remove":approval?"Existing item - approval required":"Update",Detail=detail,
            DisplayField="Email"+(Math.Max(0,before?.Emails.FindIndex(e=>e==email)??0)+1)+"DisplayName"};
    }
    internal static OutlookEntity? SelectContact(IEnumerable<OutlookEntity> matches,string corporation)
        =>matches.OrderBy(e=>e.ReadIssues.Count>0)
            .ThenByDescending(e=>e.Managed&&e.Links.Any(l=>l.Corporation==corporation))
            .ThenByDescending(e=>e.Managed).ThenByDescending(e=>e.Owned)
            .ThenBy(e=>e.Id,StringComparer.Ordinal).FirstOrDefault();

    internal static List<SyncChange> Build(string corporation,SourceTargets targets,List<OutlookEntity> existing)
    {
        var result=new List<SyncChange>();
        var unresolved=existing.Where(e=>e.Kind=="Contact"&&e.ReadIssues.Count>0).ToList();
        foreach(var (email,link) in targets.Contacts.OrderBy(x=>x.Key))
        {
            var matches=existing.Where(e=>e.Kind=="Contact"&&e.Emails.Contains(email)).ToList();
            if(matches.Count==0&&unresolved.Count>0){result.Add(new(){CanApply=false,Action="Needs review",Email=email,Detail="An existing Outlook contact has an unresolved Exchange address. A new contact is blocked until those addresses are resolved, because it could duplicate an existing contact."});continue;}
            var before=SelectContact(matches,corporation);
            if(before?.ReadIssues.Count>0){result.Add(new(){Before=before,CanApply=false,Action="Needs review",Email=email,Detail=string.Join(" ",before.ReadIssues)});continue;}
            var after=before??New("Contact");
            var links=after.Links.Where(l=>l.Corporation!=corporation).Append(link).OrderBy(l=>l.Corporation).ToList();
            var fields=new Dictionary<string,string>(after.Fields);string display=Display(links);
            fields["FullName"]=before is {Owned:false}?before.Name:display;fields["FileAs"]=display;
            int slot=before?.Emails.FindIndex(e=>e==email)??0;fields["Email"+(slot+1)+"DisplayName"]=display;
            after=after with {ManagedId=after.Managed?after.ManagedId:Guid.NewGuid().ToString("N"),Links=links,Fields=fields,Emails=before?.Emails??[email,"",""]};
            string detail=before is {Owned:false}?"Select Apply to link this existing contact. Only explicitly approved filing and email-display fields will change; other personal fields are preserved.":"Update the selected contact with all known unit associations.";
            if(matches.Count>1)detail+=$" Selected one of {matches.Count} existing contacts automatically (ID {before!.Id}); other copies are left unchanged.";
            result.Add(Change(before,after,email,detail));
        }
        // Outlook contacts can contain up to three source email addresses.
        // Plan one write per existing item, merging the source associations.
        foreach(var duplicate in result.Where(r=>r.Before is {Kind:"Contact"}&&r.After is not null).GroupBy(r=>r.Before!.Id).Where(g=>g.Count()>1).ToList())
        {
            var rows=duplicate.ToList();var before=rows[0].Before!;
            var links=rows.SelectMany(r=>r.After!.Links).GroupBy(l=>l.Corporation).OrderBy(g=>g.Key)
                .Select(g=>new ResidentLink(g.Key,g.SelectMany(l=>l.Units).Distinct().Order().ToList(),g.SelectMany(l=>l.Names).Distinct().Order().ToList())).ToList();
            string display=Display(links);var fields=new Dictionary<string,string>(before.Fields);
            if(before.Owned)fields["FullName"]=display;fields["FileAs"]=display;
            var displayFields=rows.Select(r=>r.DisplayField).Distinct().ToList();foreach(string f in displayFields)fields[f]=display;
            var after=FilingPolicy(before,rows[0].After! with{Links=links,Fields=fields})!;bool same=before.SameContent(after);
            var combined=new SyncChange{Before=before,After=after,CanApply=!same,Selected=!same&&before.Owned,Action=same?"Unchanged":before.Owned?"Update":"Existing item - approval required",Email=string.Join("; ",rows.Select(r=>r.Email)),DisplayField=displayFields[0],ExtraDisplayFields=displayFields.Skip(1).ToList(),Detail="This existing contact matches several source email addresses. One contact is preserved; approval updates both the filing name and each matching email display slot, listed in these details."};
            foreach(var row in rows)result.Remove(row);result.Add(combined);
        }
        foreach(var before in existing.Where(e=>e.Kind=="Contact"&&e.Managed&&e.Links.Any(l=>l.Corporation==corporation)&&!e.Emails.Any(targets.Contacts.ContainsKey)))
            result.Add(before.ReadIssues.Count>0?new(){Before=before,CanApply=false,Action="Needs review",Detail=string.Join(" ",before.ReadIssues)}:RemoveAssociation(before,corporation));
        foreach(var (name,members) in targets.Groups.OrderBy(x=>x.Key))
        {
            string displayName=targets.GroupNames.GetValueOrDefault(name,name);
            var matches=existing.Where(e=>e.Kind=="Group"&&(e.Name.Equals(name,StringComparison.OrdinalIgnoreCase)||e.Name.Equals(displayName,StringComparison.OrdinalIgnoreCase)||e.SourceGroup.Equals(name,StringComparison.OrdinalIgnoreCase))).ToList();
            if(matches.Count>1){result.Add(new(){CanApply=false,Action="Needs review",Detail="Duplicate Outlook group: "+name});continue;}
            var before=matches.SingleOrDefault();var after=before??New("Group");
            if(before is {Managed:true}&&(!before.Links.Any(l=>l.Corporation==corporation)||before.SourceGroup.Length>0&&!before.SourceGroup.Equals(name,StringComparison.OrdinalIgnoreCase)))
            {result.Add(new(){CanApply=false,Action="Needs review",Before=before,Detail="The requested group name is already assigned to another managed group. Choose a different name."});continue;}
            string fingerprint=Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(members.OrderBy(m=>m.Email)))));
            // Display labels are part of group identity, not just email membership.
            after=after with {SourceGroup=name,MembersFingerprint=fingerprint,ManagedId=after.Managed?after.ManagedId:Guid.NewGuid().ToString("N"),Fields=new(){{"FullName",displayName}},Members=members,Links=[new(corporation,[],[])]};
            result.Add(Change(before,after,"",$"Group: {before?.Name??"(new)"} -> {displayName}. Group membership: {before?.Members.Count??0} -> {members.Count}. Existing groups require explicit Apply approval and are never treated as app-created."));
        }
        foreach(var before in existing.Where(e=>e.Kind=="Group"&&e.Managed&&e.Links.Any(l=>l.Corporation==corporation)&&!targets.Groups.Keys.Any(n=>n.Equals(e.SourceGroup.Length>0?e.SourceGroup:e.Name,StringComparison.OrdinalIgnoreCase))))
            result.Add(Change(before,before with {Members=[]},"","This group no longer has active source members. Review before clearing it."));
        foreach(var item in unresolved.Where(e=>!result.Any(r=>r.Before?.Id==e.Id)))
            result.Add(new(){Before=item,CanApply=false,Action="Needs review",Detail="Existing Outlook contact '"+item.Name+"': "+string.Join(" ",item.ReadIssues),Email=item.Emails.FirstOrDefault(GeneratedEmailListService.IsValidEmail)??""});
        return result;
    }
    private static SyncChange RemoveAssociation(OutlookEntity before,string corporation)
    {
        var links=before.Links.Where(l=>l.Corporation!=corporation).ToList();OutlookEntity? after;
        if(before.Owned&&links.Count==0)after=null;
        else
        {
            var fields=new Dictionary<string,string>(before.Fields);
            if(before.Owned){string display=Display(links);fields["FullName"]=fields["FileAs"]=display;for(int i=0;i<3;i++)if(before.Emails[i].Length>0)fields["Email"+(i+1)+"DisplayName"]=display;}
            after=before with {Links=links,Fields=fields,ManagedId=links.Count==0?"":before.ManagedId};
        }
        var row=Change(before,after,before.Emails.FirstOrDefault(e=>e.Length>0)??"","Remove this corporation's association. Shared contacts keep other corporations; pre-existing contacts are never deleted. Removal items start unchecked.");row.Selected=false;return row;
    }

}
