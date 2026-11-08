using System.Text.Json;

namespace PropertyManager.App.OutlookSync;

internal static class FullUndoPlanner
{
    private sealed class Chain
    {
        internal List<(RecoveryRun Run,RecoveryOperation Operation)> Events=[];
        internal HashSet<string> Ids=[];
        internal HashSet<string> Markers=[];
    }
    internal static SyncPlan Build(RecoveryStore store,OutlookMailbox mailbox,string corporation,List<OutlookEntity> current)
    {
        var chains=new List<Chain>();
        foreach(var entry in store.Journal(mailbox))
        {
            var entities=new[]{entry.Operation.Before,entry.Operation.Desired,entry.Operation.Actual}.OfType<OutlookEntity>().ToList();
            var matches=chains.Where(c=>entities.Any(e=>e.Id.Length>0&&c.Ids.Contains(e.Id)||e.Managed&&c.Markers.Contains(e.ManagedId))).ToList();
            var chain=matches.FirstOrDefault()??new Chain();if(matches.Count==0)chains.Add(chain);
            foreach(var extra in matches.Skip(1)){chain.Events.AddRange(extra.Events);chain.Ids.UnionWith(extra.Ids);chain.Markers.UnionWith(extra.Markers);chains.Remove(extra);}
            chain.Events.Add(entry);foreach(var item in entities){if(item.Id.Length>0)chain.Ids.Add(item.Id);if(item.Managed)chain.Markers.Add(item.ManagedId);}
        }
        var changes=new List<SyncChange>();var seen=new HashSet<string>();
        foreach(var chain in chains)
        {
            var events=chain.Events.OrderBy(e=>e.Operation.Id).ToList();
            var scopes=events.Where(e=>e.Run.ModeIsSource()).SelectMany(e=>e.Run.Context.Corporations??[e.Run.Corporation]).Distinct().ToList();
            if(corporation!="*"&&!scopes.Contains(corporation))continue;
            var first=events[0].Operation;var last=events[^1].Operation;
            var original=first.Before;var latest=last.Actual;
            var candidates=current.Where(e=>chain.Ids.Contains(e.Id)||e.Managed&&chain.Markers.Contains(e.ManagedId)).ToList();
            foreach(var item in candidates)seen.Add(item.Id);
            var now=candidates.SingleOrDefaultIfUnique();
            string? problem=null;
            if(candidates.Count>1)problem="Multiple current items match this recovery chain. Resolve duplicates before undoing.";
            else if(now?.ReadIssues.Count>0)problem="This contact has an unresolved Exchange address. Resolve it in Outlook before undoing its changes.";
            else if(corporation!="*"&&scopes.Any(s=>s!=corporation))problem="Shared across corporations. Use all corporations to undo this item safely.";
            else if(events.Any(e=>e.Operation.Status is "Pending" or "Restoring"))problem="An interrupted operation has an uncertain outcome. Resolve it in Sync history before full undo.";
            else if(original?.Managed==true)problem="The original pre-app state is missing from retained history. Full restoration cannot be confirmed.";
            else if(original is null)
            {
                if(now is null)continue;
                if(!now.Owned||latest is null||!latest.SameVersion(now))problem="This app-created item changed after the last recorded write. It will not be deleted automatically.";
            }
            else if(now is null)problem="The existing personal item is missing or moved. It will not be recreated automatically.";
            else if(latest is null)problem="The latest saved state is unavailable. Review the recovery history.";
            OutlookEntity? target=null;
            if(problem is null&&original is not null&&now is not null&&latest is not null)
            {
                if(original.Kind=="Group")
                {
                    if(!latest.SameVersion(now))problem="This group changed after the app last saved it. Membership and name were left unchanged.";
                    else target=original with{Id=now.Id,Modified=now.Modified};
                }
                else
                {
                    if(!now.Emails.SequenceEqual(latest.Emails))problem="Email addresses changed after the app last saved this contact. Review the display labels before undoing.";
                    var fields=new Dictionary<string,string>(now.Fields);
                    // Only undo fields actually changed by app writes. Preserve unrelated user edits.
                    foreach(string key in new[]{"FullName","FileAs","Email1DisplayName","Email2DisplayName","Email3DisplayName"})
                    {
                        var writes=events.Where(e=>e.Operation.Before is not null&&e.Operation.Actual is not null&&e.Operation.Before.Fields.GetValueOrDefault(key,"")!=e.Operation.Actual.Fields.GetValueOrDefault(key,"")).ToList();
                        if(writes.Count==0)continue;
                        string baseline=writes[0].Operation.Before!.Fields.GetValueOrDefault(key,"");
                        string expected=latest.Fields.GetValueOrDefault(key,"");
                        if(now.Fields.GetValueOrDefault(key,"")!=expected){problem="The "+key+" field changed after the last app write. Review it before undoing.";break;}
                        fields[key]=baseline;
                    }
                    if(problem is null&&(now.Owned!=latest.Owned||now.SourceGroup!=latest.SourceGroup||now.ManagedId!=latest.ManagedId||JsonSerializer.Serialize(now.Links)!=JsonSerializer.Serialize(latest.Links)))problem="App associations changed after the last recorded write. Review this contact.";
                    bool formatCaptured=now.FilingFormatCaptured;int? format=now.FilingFormat;
                    var formatWrites=events.Where(e=>e.Operation.Before is {FilingFormatCaptured:true}&&e.Operation.Actual is {FilingFormatCaptured:true}&&e.Operation.Before.FilingFormat!=e.Operation.Actual.FilingFormat).ToList();
                    if(formatWrites.Count>0)
                    {
                        if(!now.FilingFormatCaptured||!latest.FilingFormatCaptured||now.FilingFormat!=latest.FilingFormat)problem="The filing-name format changed after the last app write. Review it before undoing.";
                        else{formatCaptured=true;format=formatWrites[0].Operation.Before!.FilingFormat;}
                    }
                    if(problem is null)target=now with{Fields=fields,FilingFormatCaptured=formatCaptured,FilingFormat=format,ManagedId=original.ManagedId,Owned=original.Owned,Links=original.Links,SourceGroup=original.SourceGroup};
                }
            }
            if(problem is null&&now is not null&&target is not null&&now.SameContent(target))continue;
            changes.Add(new(){Before=now,After=problem is null?target:null,RestoreOriginal=true,CanApply=problem is null,Selected=false,NeedsManualSelection=true,
                Action=problem is not null?"Needs review":target is null?"Remove":"Restore original",Email=(now??latest??original)?.Emails.FirstOrDefault(e=>e.Length>0)??"",
                Detail=problem??(target is null?"Remove the unchanged app-created item.":"Restore original app-modified fields and associations. Preserve unrelated contact data.")});
        }
        foreach(var item in current.Where(e=>e.Managed&&!seen.Contains(e.Id)&&(corporation=="*"||e.Links.Any(l=>l.Corporation==corporation))))
            changes.Add(new(){Before=item,CanApply=false,Action="Needs review",Detail="No matching recovery history exists on this Windows account. The original state cannot be confirmed.",Email=item.Emails.FirstOrDefault()??""});
        return new(mailbox,corporation,"Undo all",changes);
    }
    private static bool ModeIsSource(this RecoveryRun run)=>run.Context.Mode is not "Undo all" and not "Checkpoint restore";
    private static OutlookEntity? SingleOrDefaultIfUnique(this List<OutlookEntity> values)=>values.Count==1?values[0]:null;
}
