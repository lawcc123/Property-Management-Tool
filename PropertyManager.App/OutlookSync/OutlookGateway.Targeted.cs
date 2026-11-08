namespace PropertyManager.App.OutlookSync;

internal sealed partial class OutlookGateway
{
    public List<OutlookEntity> ReadForChanges(OutlookMailbox mailbox,
        IReadOnlyList<(OutlookEntity? Before,OutlookEntity? After)> changes,
        Action<int,int>? progress=null,CancellationToken cancellation=default)
    {
        dynamic? app=null,session=null,folder=null,items=null;
        var found=new Dictionary<string,OutlookEntity>();
        var before=changes.Where(c=>c.Before is not null).Select(c=>c.Before!).DistinctBy(e=>e.Id).ToList();
        var newEmails=changes.Where(c=>c.Before is null&&c.After?.Kind=="Contact")
            .SelectMany(c=>c.After!.Emails).Where(e=>e.Length>0).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var groupNames=changes.Where(c=>c.After?.Kind=="Group").Select(c=>c.After!.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        try
        {
            cancellation.ThrowIfCancellationRequested();app=Connect();
            session=ReadValue<object>("open the Outlook MAPI session",()=>app.GetNamespace("MAPI"));
            folder=ReadValue<object>("open the selected Contacts folder",()=>session.GetFolderFromID(mailbox.FolderId,mailbox.StoreId));
            // Fully re-read each changed item by ID, including labels, addresses,
            // tracking and group membership. Never trust cached values for writes.
            for(int i=0;i<before.Count;i++)
            {
                cancellation.ThrowIfCancellationRequested();progress?.Invoke(i+1,before.Count);
                dynamic? item=null,parent=null;
                try
                {
                    item=ReadValue<object>("reopen changed item '"+before[i].Name+"'",()=>session.GetItemFromID(before[i].Id,mailbox.StoreId));
                    parent=ReadValue<object>("check the changed item's folder",()=>item.Parent);
                    if(ReadValue<string>("read the changed item's folder ID",()=>Text(parent.EntryID))!=mailbox.FolderId)
                        throw new InvalidOperationException("A selected Outlook item moved out of this Contacts folder. Generate a new preview; no changes were written.");
                    OutlookEntity fresh=ReadItem(item);
                    if(fresh.Id!=before[i].Id)throw new InvalidOperationException("Outlook returned another item for a selected ID. No changes were written.");
                    found[fresh.Id]=fresh;
                }
                finally{Release(parent);Release(item);}
            }
            // Duplicate protection still examines the current folder, but only
            // address identities / group names of unrelated items, not their
            // labels, markers, filing rules or every group's members.
            if(newEmails.Count>0||groupNames.Count>0)
                items=ReadValue<object>("open Contacts for duplicate checks",()=>folder.Items);
            if(newEmails.Count>0)
                ScanIdentities((object)items!,"IPM.Contact",item=>
                {
                    string[] addresses=[ReadValue<string>("read Email1 identity",()=>Text(item.Email1Address)).Trim().ToLowerInvariant(),
                        ReadValue<string>("read Email2 identity",()=>Text(item.Email2Address)).Trim().ToLowerInvariant(),
                        ReadValue<string>("read Email3 identity",()=>Text(item.Email3Address)).Trim().ToLowerInvariant()];
                    bool relevant=addresses.Any(newEmails.Contains),exchange=false;
                    for(int slot=0;slot<3&&!relevant;slot++)
                    {
                        if(addresses[slot].Length==0||ValidSmtp(addresses[slot]))continue;
                        string type=slot==0?ReadValue<string>("read Email1 identity type",()=>Text(item.Email1AddressType)):
                            slot==1?ReadValue<string>("read Email2 identity type",()=>Text(item.Email2AddressType)):
                                    ReadValue<string>("read Email3 identity type",()=>Text(item.Email3AddressType));
                        if(type.Equals("EX",StringComparison.OrdinalIgnoreCase))exchange=true;
                    }
                    if(!relevant&&!exchange)return;
                    OutlookEntity entity=ReadItem(item);
                    if(entity.ReadIssues.Count>0||entity.Emails.Any(newEmails.Contains))found[entity.Id]=entity;
                },found,progress,cancellation);
            if(groupNames.Count>0)
                ScanIdentities((object)items!,"IPM.DistList",item=>
                {
                    if(groupNames.Contains(ReadValue<string>("read group identity",()=>Text(item.DLName))))
                    {OutlookEntity entity=ReadItem(item);found[entity.Id]=entity;}
                },found,progress,cancellation);
            return found.Values.ToList();
        }
        finally{Release(items);Release(folder);Release(session);Release(app);}
    }
    private static void ScanIdentities(dynamic items,string messageClass,Action<dynamic> inspect,
        Dictionary<string,OutlookEntity> found,Action<int,int>? progress,CancellationToken cancellation)
    {
        dynamic? filtered=null;
        try
        {
            // Fixed built-in message classes only: no user-entered filter text.
            filtered=ReadValue<object>("filter Contacts identities",()=>items.Restrict("[MessageClass] = '"+messageClass+"'"));
            int total=ReadValue<int>("count identity candidates",()=>(int)filtered.Count);
            for(int i=1;i<=total;i++)
            {
                cancellation.ThrowIfCancellationRequested();progress?.Invoke(i,total);dynamic? item=null;
                try
                {
                    item=ReadValue<object>("open identity candidate "+i+" of "+total,()=>filtered.Item(i));
                    string id=ReadValue<string>("read identity item ID",()=>Text(item.EntryID));
                    if(!found.ContainsKey(id))inspect(item);
                }
                catch(OperationCanceledException){throw;}
                catch(Exception ex){throw new InvalidOperationException("Checking "+messageClass+" identity "+i+" of "+total+": "+ex.Message,ex);}
                finally{Release(item);}
            }
        }
        finally{Release(filtered);}
    }
}
