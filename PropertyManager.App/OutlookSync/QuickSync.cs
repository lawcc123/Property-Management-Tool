using System.IO;
using System.Text.Json;
using PropertyManager.App.EmailLists;

namespace PropertyManager.App.OutlookSync;

internal sealed class QuickSyncSettings
{
    public OutlookMailbox? Destination {get;set;}
    public List<string> ExcludedCorporations {get;set;}=[];
    internal static string SettingsPath=>Path.Combine(Services.AppSettings.DataDirectory,"OutlookSync","quick-sync.json");
    internal static QuickSyncSettings Load(string? path=null)=>Settings.SettingsDatabase.Read(path??SettingsPath) is {} json?JsonSerializer.Deserialize<QuickSyncSettings>(json)??throw new InvalidDataException("Quick sync settings are empty."):new();
    internal void Save(string? path=null)
    {
        path??=SettingsPath;Settings.SettingsDatabase.Write(path,JsonSerializer.Serialize(this,new JsonSerializerOptions{WriteIndented=true}));
    }
    internal OutlookMailbox Resolve(IEnumerable<OutlookMailbox> available)=>Destination is null
        ?throw new InvalidOperationException("Connect to Outlook, select your working Contacts folder, and click Save destination first.")
        :available.SingleOrDefault(b=>b.StoreId==Destination.StoreId&&b.FolderId==Destination.FolderId)
          ??throw new InvalidOperationException("Your saved Contacts destination is unavailable. Select and save the intended folder again. No archive or replacement folder was selected.");
}
internal sealed record QuickSyncSource(string Corporation,SourceTargets Targets,string? Path=null,string? Hash=null);
internal sealed record QuickSyncReview(SyncPlan Plan)
{
    public string Label=>Plan.Corporation+" - "+Plan.Changes.Count(c=>c.Action!="Unchanged")+" remaining";
}
internal sealed record QuickSyncResult(List<QuickSyncReview> Reviews,int Created,int Updated,int Removed,int Unchanged,List<string> Warnings);

internal sealed class QuickSyncService(IOutlookGateway gateway,RecoveryStore recovery)
{
    // A managed, linked personal item was explicitly linked by an earlier Apply.
    // Ownership is never promoted: personal contacts still cannot be deleted.
    internal static bool Routine(SyncChange row,string corporation)=>row.CanApply&&row.Action!="Unchanged"&&
        !(row.Before?.Kind=="Group"&&row.After is null)&&
        (row.Before is null||row.Before.Owned||row.Before.Managed&&row.Before.Links.Any(l=>l.Corporation==corporation));

    internal static List<SyncChange> Build(QuickSyncSource source,IReadOnlyList<QuickSyncSource> sources,List<OutlookEntity> existing)
    {
        var rows=SyncPlanner.Build(source.Corporation,source.Targets,existing);
        // A resident may move to another selected corporation before that
        // corporation has ever been synced. Keep the same Outlook contact.
        for(int i=0;i<rows.Count;i++)
        {
            var row=rows[i];if(!row.CanApply||row.Before is not {Kind:"Contact",Owned:true} before||row.After is not null)continue;
            var otherSources=sources.Where(s=>s.Corporation!=source.Corporation).ToList();
            var links=otherSources.SelectMany(s=>before.Emails.Where(e=>s.Targets.Contacts.ContainsKey(e)&&
                    SyncPlanner.SelectContact(existing.Where(c=>c.Kind=="Contact"&&c.Emails.Contains(e)),s.Corporation)?.Id==before.Id).Select(e=>s.Targets.Contacts[e]))
                .GroupBy(l=>l.Corporation).Select(g=>new ResidentLink(g.Key,g.SelectMany(l=>l.Units).Distinct().Order().ToList(),g.SelectMany(l=>l.Names).Distinct().Order().ToList())).ToList();
            if(links.Count==0)
            {
                if(otherSources.Any(s=>before.Emails.Any(s.Targets.Contacts.ContainsKey)))
                    rows[i]=new(){Before=before,After=before,CanApply=false,Action="Unchanged",Email=row.Email,Detail="Another existing contact is selected for this active email. This duplicate is left unchanged."};
                continue;
            }
            string label=SyncPlanner.Display(links);var fields=new Dictionary<string,string>(before.Fields){["FullName"]=label,["FileAs"]=label};
            for(int n=0;n<before.Emails.Count;n++)if(before.Emails[n].Length>0)fields["Email"+(n+1)+"DisplayName"]=label;
            rows[i]=new(){Before=before,After=before with{Links=links,Fields=fields,FilingFormat=before.FilingFormatCaptured?-1:before.FilingFormat},Email=row.Email,Action="Update",Detail="Keep this contact: its email is active in another selected corporation. Update its unit associations."};
        }
        foreach(var row in rows)row.Selected=Routine(row,source.Corporation);
        return rows;
    }
    internal List<string> LastApplyWarnings {get;}=[];
    internal static SyncPlan CombineReviews(OutlookMailbox destination,QuickSyncResult result)
    {
        var entries=result.Reviews.SelectMany(review=>review.Plan.Changes.Select(row=>(Corporation:review.Plan.Corporation,Row:row))).ToList();
        foreach(var entry in entries)entry.Row.Corporation=entry.Corporation;
        var combined=entries.Select(e=>e.Row).ToList();
        // A shared personal contact may need first-time approval in both sources.
        // Show and write it once, retaining both corporations and each email slot.
        foreach(var duplicate in entries.Where(e=>e.Row.Before is {Kind:"Contact"}&&e.Row.After is not null).GroupBy(e=>e.Row.Before!.Id).Where(g=>g.Count()>1))
        {
            var group=duplicate.ToList();var before=group[0].Row.Before!;
            if(group.Any(e=>!e.Row.CanApply||!before.SameSyncState(e.Row.Before)))
            {
                foreach(var entry in group)combined.Remove(entry.Row);
                combined.Add(new(){Before=before,Action="Needs review",CanApply=false,Corporation=string.Join("; ",group.Select(e=>e.Corporation).Distinct()),Detail="This shared contact has conflicting proposals. Run Update all corporations again after resolving its Outlook details."});continue;
            }
            var scopes=group.Select(e=>e.Corporation).ToHashSet();
            var links=before.Links.Where(l=>!scopes.Contains(l.Corporation))
                .Concat(group.SelectMany(e=>e.Row.After!.Links.Where(l=>l.Corporation==e.Corporation)))
                .GroupBy(l=>l.Corporation).OrderBy(g=>g.Key)
                .Select(g=>new ResidentLink(g.Key,g.SelectMany(l=>l.Units).Distinct().Order().ToList(),g.SelectMany(l=>l.Names).Distinct().Order().ToList())).ToList();
            string label=SyncPlanner.Display(links);var fields=new Dictionary<string,string>(before.Fields){["FileAs"]=label};
            if(before.Owned)fields["FullName"]=label;
            var displays=group.SelectMany(e=>e.Row.DisplayFields).Distinct().ToList();foreach(string display in displays)fields[display]=label;
            var after=group[0].Row.After! with{Links=links,Fields=fields};
            foreach(var entry in group)combined.Remove(entry.Row);
            combined.Add(new(){Before=before,After=after,Corporation=string.Join("; ",scopes.Order()),
                Action=before.Owned?"Update":"Existing item - approval required",Selected=group.All(e=>e.Row.Selected),
                Email=string.Join("; ",group.Select(e=>e.Row.Email).Distinct()),DisplayField=displays[0],ExtraDisplayFields=displays.Skip(1).ToList(),
                Detail="One shared contact for both corporations. Approval updates the filing and email display labels with all unit associations; other personal fields are preserved."});
        }
        return new(destination,"All","Sync",combined);
    }
    internal string ApplyReviews(SyncPlan plan,IReadOnlyList<QuickSyncSource> sources,Action<string,int,int> progress,CancellationToken token)
    {
        LastApplyWarnings.Clear();
        foreach(var source in sources)if(source.Path is not null&&SyncEngine.Hash(source.Path)!=source.Hash)
            throw new InvalidOperationException(source.Corporation+": the source workbook changed. Run Update all corporations again before applying approvals.");
        token.ThrowIfCancellationRequested();
        string run=recovery.Start(plan,sources.Select(s=>s.Corporation).ToList());
        try
        {
            var engine=new SyncEngine(gateway,recovery);engine.Apply(plan,progress,token,run);
            LastApplyWarnings.AddRange(engine.LastApplyWarnings);
            recovery.Finish(run,LastApplyWarnings.Count==0?"Completed":"Completed - "+LastApplyWarnings.Count+" item(s) need review");return run;
        }
        catch(OperationCanceledException){recovery.Finish(run,"Cancelled - may be partial");throw;}
        catch{recovery.Finish(run,"Failed - may be partial");throw;}
    }
    internal QuickSyncResult Run(OutlookMailbox destination,IReadOnlyList<QuickSyncSource> sources,Action<string,int,int> progress,CancellationToken token)
    {
        if(sources.Count==0)throw new InvalidOperationException("Select at least one corporation in Update settings.");
        if(sources.Select(s=>s.Corporation).Distinct().Count()!=sources.Count)throw new InvalidOperationException("More than one source folder is configured for the same corporation. Resolve it in Manage folders.");
        // Validate every workbook before the first write.
        foreach(var source in sources)if(source.Path is not null&&SyncEngine.Hash(source.Path)!=source.Hash)throw new InvalidOperationException(source.Corporation+": the workbook changed while preparing this update. Try again.");
        var groupNames=sources.SelectMany(s=>s.Targets.Groups.Keys.Select(k=>(Corporation:s.Corporation,Name:s.Targets.GroupNames.GetValueOrDefault(k,k)))).ToList();
        if(groupNames.GroupBy(g=>g.Name,StringComparer.OrdinalIgnoreCase).Any(g=>g.Count()>1))throw new InvalidOperationException("Two source groups use the same Outlook group name. Choose unique names in Group settings before updating.");
        if(gateway is OutlookGateway live)live.TargetGroupNames=groupNames.Select(g=>g.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        int created=0,updated=0,removed=0;var warnings=new List<string>();
        token.ThrowIfCancellationRequested();
        string batchRun=recovery.Start(new(destination,sources.Count>1?"All":sources[0].Corporation,"Update all corporations",[]),sources.Select(s=>s.Corporation).ToList());
        try
        {
            var current=gateway.Read(destination,(d,t)=>progress("Reading Outlook for all selected corporations",d,t),token);
            foreach(var source in sources)
            {
                token.ThrowIfCancellationRequested();
                var rows=Build(source,sources,current);var selected=rows.Where(c=>c.Selected).ToList();
                if(selected.Count==0)continue;
                var engine=new SyncEngine(gateway,recovery);
                engine.Apply(new(destination,source.Corporation,"Sync",rows,source.Path,source.Hash),(m,d,t)=>progress(source.Corporation+" - "+m,d,t),token,batchRun);
                // Use Outlook's actual readback, including provider changes, for the
                // next corporation. This cache lasts only for this batch.
                foreach(var saved in engine.SavedChanges)
                {
                    if(saved.Before is not null)current.RemoveAll(e=>e.Id==saved.Before.Id);
                    if(saved.Actual is not null){current.RemoveAll(e=>e.Id==saved.Actual.Id);current.Add(saved.Actual);}
                }
                created+=selected.Count(c=>c.Before is null);removed+=selected.Count(c=>c.After is null);updated+=selected.Count(c=>c.Before is not null&&c.After is not null);
                warnings.AddRange(engine.LastApplyWarnings);
            }
            var result=Review(destination,sources,progress,token);
            int needsReview=result.Reviews.Sum(r=>r.Plan.Changes.Count(c=>c.Action!="Unchanged"));
            recovery.Finish(batchRun,needsReview==0&&warnings.Count==0?"Completed":$"Completed - {needsReview} item(s) need review"+(warnings.Count>0?$"; {warnings.Count} save warning(s)":""));
            return result with{Created=created,Updated=updated,Removed=removed,Warnings=warnings};
        }
        catch(OperationCanceledException){recovery.Finish(batchRun,"Cancelled - may be partial");throw;}
        catch{recovery.Finish(batchRun,"Failed - may be partial");throw;}
    }
    internal QuickSyncResult Review(OutlookMailbox destination,IReadOnlyList<QuickSyncSource> sources,Action<string,int,int> progress,CancellationToken token)
    {
        foreach(var source in sources)if(source.Path is not null&&SyncEngine.Hash(source.Path)!=source.Hash)throw new InvalidOperationException(source.Corporation+": the source workbook changed. Run Update all again to load the latest data.");
        // Re-read after all writes so shared contacts and pending approvals
        // reflect Outlook's saved state, including after a manual approval.
        var final=gateway.Read(destination,(d,t)=>progress("Checking final saved results",d,t),token);var reviews=new List<QuickSyncReview>();int unchanged=0;
        foreach(var source in sources)
        {
            var rows=Build(source,sources,final);unchanged+=rows.Count(c=>c.Action=="Unchanged");
            if(rows.Any(c=>c.Action!="Unchanged"))reviews.Add(new(new(destination,source.Corporation,"Sync",rows.Where(c=>c.Action!="Unchanged").ToList(),source.Path,source.Hash)));
        }
        return new(reviews,0,0,0,unchanged,[]);
    }
}

internal static class QuickSyncWorkbook
{
    internal static int SheetIndex(IEnumerable<string> names)
    {
        var sheets=names.ToList();if(sheets.Count==1)return 0;
        var matches=sheets.Select((name,index)=>(name,index)).Where(s=>System.Text.RegularExpressions.Regex.Replace(s.name,"[^a-zA-Z]","").ToLowerInvariant() is "occupantsdata" or "occupantdata").ToList();
        if(matches.Count!=1)throw new InvalidDataException("The Occupant Data worksheet could not be identified uniquely. Use a standard BuildingLink Occupant Data export before running Update all.");
        return matches[0].index;
    }
    internal static Task<RecipientFileData> LoadAsync(string path)=>Task.Run(()=>
    {
        if(Path.GetExtension(path).Equals(".xls",StringComparison.OrdinalIgnoreCase))
        {
            var sheets=LegacyRecipientReader.Read(path);return sheets[SheetIndex(sheets.Select(s=>s.Sheet))].Data;
        }
        var modern=ExcelRecipientReader.ListWorksheets(path);return ExcelRecipientReader.ReadSheet(path,modern[SheetIndex(modern.Select(s=>s.Name))]);
    });
}
