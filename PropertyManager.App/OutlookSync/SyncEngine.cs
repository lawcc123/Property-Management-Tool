using System.IO;
using System.Security.Cryptography;
using System.Text.Json;

namespace PropertyManager.App.OutlookSync;

internal sealed class SyncEngine(IOutlookGateway gateway,RecoveryStore recovery)
{
    internal List<string> LastApplyWarnings {get;}=[];
    internal List<(OutlookEntity? Before,OutlookEntity? Actual)> SavedChanges {get;}=[];
    internal static string Hash(string path){using var file=File.OpenRead(path);return Convert.ToHexString(SHA256.HashData(file));}
    private FileStream Lock()=>new(Path.Combine(recovery.DirectoryPath,"sync.lock"),FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);
    internal string Apply(SyncPlan plan,Action<string,int,int> progress,CancellationToken cancellation,string? batchRun=null)
    {
        LastApplyWarnings.Clear();SavedChanges.Clear();
        using var held=Lock();
        if(plan.SourcePath is not null&&Hash(plan.SourcePath)!=plan.SourceHash)throw new InvalidOperationException("The source workbook changed after preview. Generate a new preview.");
        var changes=plan.Changes.Where(r=>r.Selected&&r.CanApply).Select(r=>(r.Before,After:r.ApprovedAfter())).Where(r=>r.Before is null||!r.Before.SameContent(r.After)).ToList();
        if(changes.Count==0)throw new InvalidOperationException("Select at least one changed item. Existing personal contacts need explicit field approvals for filing and email display names.");
        var current=plan.Mode=="Sync"&&gateway is IOutlookTargetedGateway targeted
            ?targeted.ReadForChanges(plan.Mailbox,changes,(d,t)=>progress("Checking selected changes and duplicate contacts",d,t),cancellation)
            :gateway.Read(plan.Mailbox,(d,t)=>progress("Checking Outlook contacts",d,t),cancellation);
        for(int index=0;index<changes.Count;index++)
        {
            var (before,after)=changes[index];
            if(before?.ReadIssues.Count>0||after?.ReadIssues.Count>0)throw new InvalidOperationException("A selected contact has an unresolved email address. Resolve it in Outlook and generate a new preview.");
            if(before is null&&after?.Kind=="Contact"&&current.Any(e=>e.Kind=="Contact"&&e.ReadIssues.Count>0))throw new InvalidOperationException("An unresolved existing Outlook address prevents ruling out duplicate contacts. No contacts were created; resolve the flagged addresses and preview again.");
            if(before is not null)
            {
                var fresh=current.SingleOrDefault(e=>e.Id==before.Id);
                if(!before.SameVersion(fresh))
                {
                    // Sync writes only names, email labels, memberships and markers.
                    // A server timestamp alone is not a change to those values. Rebase
                    // the recovery before-image, but keep undo/removal checks strict.
                    if(plan.Mode=="Sync"&&after is not null&&before.SameSyncState(fresh))
                    {before=fresh!;changes[index]=(before,after);}
                    else throw new InvalidOperationException("Outlook changed after preview for "+before.Kind+" '"+before.Name+"': "+string.Join(", ",before.ChangedFields(fresh))+". Generate a new preview. Nothing in this run was written.");
                }
            }
            if(after?.Kind=="Group"&&current.Any(e=>e.Kind=="Group"&&e.Id!=before?.Id&&e.Name.Equals(after.Name,StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("Another Outlook group now uses the requested name. Generate a new preview before applying a rename.");
            if(before is null&&after is not null&&current.Any(e=>e.Kind==after.Kind&&(after.Kind=="Group"?e.Name.Equals(after.Name,StringComparison.OrdinalIgnoreCase):e.Emails.Any(v=>v.Length>0&&after.Emails.Contains(v)))))
                throw new InvalidOperationException("A matching Outlook item appeared after preview. Generate a new preview to avoid duplicates.");
        }
        cancellation.ThrowIfCancellationRequested();string run=batchRun??recovery.Start(plan);int completed=0;
        try
        {
            foreach(var (before,after) in changes)
            {
                cancellation.ThrowIfCancellationRequested();progress("Saving recovery record",completed,changes.Count);
                var backedBefore=before is not null&&(after is null||before.Kind=="Group")?gateway.CaptureBackup(plan.Mailbox,before):before;
                long operation=recovery.Prepare(run,backedBefore,after);
                progress((before is null?"Creating ":after is null?"Removing ":"Updating ")+(after??before)!.Kind,completed,changes.Count);
                var actual=gateway.Write(plan.Mailbox,before,after);recovery.Applied(operation,actual);SavedChanges.Add((before,actual));completed++;
                var differences=OutlookWriteVerification.Differences(after,actual);
                if(differences.Count>0)LastApplyWarnings.Add((after??before)!.Kind+" '"+(after??before)!.Name+"': Outlook did not retain "+string.Join(", ",differences)+". See the desired and actual values in this run's recovery details.");
                progress("Outlook item saved",completed,changes.Count);
            }
            if(batchRun is null)recovery.Finish(run,LastApplyWarnings.Count==0?"Completed":"Completed - "+LastApplyWarnings.Count+" item(s) need review");recovery.Retain();return run;
        }
        catch(OperationCanceledException){recovery.Finish(run,"Cancelled - may be partial");throw;}
        catch{recovery.Finish(run,"Failed - may be partial");throw;}
    }
    internal List<string> Restore(RecoveryRun run,Action<string,int,int> progress,CancellationToken cancellation)
    {
        using var held=Lock();var operations=recovery.Operations(run.Id);var conflicts=new List<string>();var restoredStates=new Dictionary<string,OutlookEntity>();int completed=0;
        recovery.Finish(run.Id,"Restoring");
        try
        {
            var items=gateway.Read(run.Context.Mailbox,(d,t)=>progress("Checking Outlook contacts",d,t),cancellation);
            foreach(var operation in operations)
            {
                cancellation.ThrowIfCancellationRequested();progress("Checking recovery item",completed++,operations.Count);
                if(operation.Status=="Restored")continue;
                OutlookEntity? Find(OutlookEntity? target)=>target is null?null:items.SingleOrDefault(e=>e.Id==target.Id&&target.Id.Length>0)??(target.Managed?items.SingleOrDefault(e=>e.ManagedId==target.ManagedId):null);
                var current=Find(operation.Actual??operation.Desired??operation.Before);
                // A crash during restore may already have saved the original content.
                if(operation.Status=="Restoring"&&(operation.Before is null?current is null:operation.Before.SameContent(current)))
                {recovery.OperationStatus(operation.Id,"Restored");continue;}
                bool unchanged=operation.Status is "Pending" or "Restoring"
                    ? operation.Desired is null?current is null:operation.Desired.SameContent(current)
                    : operation.Actual is null?current is null:operation.Actual.SameVersion(current);
                // A previous step in this same restore may have reverted a shared
                // contact to this operation's content with a fresh Outlook timestamp.
                // Only trust the exact state returned by that verified restore write.
                if(!unchanged && current is not null && operation.Actual is not null &&
                    restoredStates.TryGetValue(current.Id,out var restoredState) &&
                    restoredState.SameVersion(current) && operation.Actual.SameContent(current))unchanged=true;
                if(!unchanged)
                {
                    if(operation.Status=="Pending"&&(operation.Before is null?current is null:operation.Before.SameVersion(current)))
                    {recovery.OperationStatus(operation.Id,"Restored");continue;}
                    conflicts.Add((operation.Before??operation.Desired)?.Name??"Unknown item");continue;
                }
                recovery.OperationStatus(operation.Id,"Restoring");
                var restore=operation.Before;
                if(restore is not null&&current is null)
                {
                    if(items.Any(e=>e.Kind==restore.Kind&&(e.Kind=="Group"?e.Name.Equals(restore.Name,StringComparison.OrdinalIgnoreCase):e.Emails.Any(v=>v.Length>0&&restore.Emails.Contains(v)))))
                    {conflicts.Add(restore.Name+" - an existing item now matches this restored record");continue;}
                    if(restore.FullItemBackup is null)throw new InvalidOperationException("The full deleted-item backup is missing; restoration was stopped.");
                    restore=restore with{Id="",Modified=""};
                }
                // Journal checkpoint restoration too, so full undo follows the actual write chain.
                string restoreRun=recovery.Start(new(run.Context.Mailbox,run.Context.Corporation,"Checkpoint restore",[]));
                var restoreBefore=current is not null&&(restore is null||current.Kind=="Group")?gateway.CaptureBackup(run.Context.Mailbox,current):current;
                long restoreOperation=recovery.Prepare(restoreRun,restoreBefore,restore);
                var actual=gateway.Write(run.Context.Mailbox,current,restore);
                recovery.Applied(restoreOperation,actual);recovery.Finish(restoreRun,"Completed");
                if(current is not null)items.RemoveAll(e=>e.Id==current.Id);
                if(actual is not null){items.Add(actual);restoredStates[actual.Id]=actual;}
                if(current is not null&&actual is null)restoredStates.Remove(current.Id);
                recovery.OperationStatus(operation.Id,"Restored");
            }
            recovery.Finish(run.Id,conflicts.Count==0?"Restored":"Restore conflicts");recovery.Retain();return conflicts;
        }
        catch(OperationCanceledException){recovery.Finish(run.Id,"Restore interrupted");throw;}
        catch{recovery.Finish(run.Id,"Restore interrupted");throw;}
    }
    internal string Report(RecoveryRun run)=>JsonSerializer.Serialize(new{run,operations=recovery.Operations(run.Id).Select(o=>new{o.Id,o.Status,HasFullItemBackup=o.Before?.FullItemBackup is not null,Before=o.Before is null?null:o.Before with{FullItemBackup=null},Desired=o.Desired is null?null:o.Desired with{FullItemBackup=null},Actual=o.Actual is null?null:o.Actual with{FullItemBackup=null}})},new JsonSerializerOptions{WriteIndented=true});
}
