using System.IO;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace PropertyManager.App.OutlookSync;

// Created and called only on this dedicated STA thread. No mail is ever sent.
internal sealed class OutlookStaWorker : IDisposable
{
    private readonly BlockingCollection<Action> work=new();
    internal OutlookStaWorker(){var thread=new Thread(()=>{foreach(var action in work.GetConsumingEnumerable())action();}){IsBackground=true,Name="PropertyManager Outlook STA"};thread.SetApartmentState(ApartmentState.STA);thread.Start();}
    internal Task<T> Run<T>(Func<T> action)
    {
        var tcs=new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        work.Add(()=>{try{tcs.SetResult(action());}catch(Exception ex){tcs.SetException(ex);}});return tcs.Task;
    }
    public void Dispose()=>work.CompleteAdding();
}
internal sealed record OutlookMarker(string App,string Id,bool Owned,List<ResidentLink> Links,string SourceGroup="",string MembersFingerprint="");
internal sealed class OutlookReadException : InvalidOperationException
{
    internal OutlookReadException(string step,Exception inner):base("Outlook could not "+step+" (0x"+Code(inner).ToString("X8")+"). "+inner.Message,inner){HResult=Code(inner);}
    private static int Code(Exception error)
    {
        for(Exception? current=error;current is not null;current=current.InnerException)if(current is COMException)return current.HResult;
        return error.HResult;
    }
}
internal sealed partial class OutlookGateway : IOutlookGateway,IOutlookTargetedGateway
{
    internal HashSet<string> TargetGroupNames {get;set;}=new(StringComparer.OrdinalIgnoreCase);
    internal bool ShouldReadGroup(string name,bool marked)=>marked||name.StartsWith("DEMO ",StringComparison.OrdinalIgnoreCase)||TargetGroupNames.Contains(name);
    private readonly Func<object>? applicationFactory;
    internal string ProfileDiagnostics {get;private set;}="Not captured";
    internal OutlookGateway(Func<object>? applicationFactory=null)=>this.applicationFactory=applicationFactory;
    private static T ReadValue<T>(string step,Func<T> get)
    {
        for(int attempt=0;;attempt++)
        {
            try{return get();}
            catch(COMException ex)when(attempt<2&&(ex.ErrorCode==unchecked((int)0x80010001)||ex.ErrorCode==unchecked((int)0x8001010A)))
            {Thread.Sleep(150*(attempt+1));}
            catch(COMException ex){throw new OutlookReadException(step,ex);}
        }
    }

    private const string Property="http://schemas.microsoft.com/mapi/string/{5DFEDCD5-F0DD-58CB-A911-F2EBFB284159}/PropertyManagerContactSync";
    private const string Owner="PropertyManagerDemo.Outlook.v1";
    private static void Release(object? value){if(value is not null&&Marshal.IsComObject(value))Marshal.ReleaseComObject(value);}
    private static string Text(object? value)=>Convert.ToString(value,System.Globalization.CultureInfo.InvariantCulture)??"";
    private dynamic Connect()
    {
        PropertyManager.App.PublicEdition.BlockLiveOperation();
        if(applicationFactory is not null)return applicationFactory();
        var type=Type.GetTypeFromProgID("Outlook.Application")??throw new InvalidOperationException("Classic Outlook is required. Configure its profile, then try again. New Outlook is not supported by this integration.");
        return Activator.CreateInstance(type)??throw new InvalidOperationException("Classic Outlook could not be opened.");
    }
    public List<OutlookMailbox> Mailboxes()
    {
        dynamic app=Connect(),session=app.GetNamespace("MAPI"),stores=session.Stores;var result=new List<OutlookMailbox>();
        try
        {
            for(int i=1;i<=(int)stores.Count;i++)
            {
                dynamic? store=null,folder=null;
                try{store=stores.Item(i);folder=store.GetDefaultFolder(10);result.Add(new(Text(store.StoreID),Text(folder.EntryID),Text(store.DisplayName)+" - "+Text(folder.FolderPath)));}
                catch(COMException){}finally{Release(folder);Release(store);}
            }
            if(result.Count==0)throw new InvalidOperationException("No accessible Classic Outlook Contacts folder was found.");return result;
        }
        finally{Release(stores);Release(session);Release(app);}
    }
    public List<OutlookEntity> Read(OutlookMailbox mailbox,Action<int,int>? progress=null,CancellationToken cancellation=default)
    {
        dynamic? app=null,session=null,folder=null,items=null;
        var result=new List<OutlookEntity>();
        try
        {
            app=Connect();session=ReadValue<object>("open the Outlook MAPI session",()=>app.GetNamespace("MAPI"));
            folder=ReadValue<object>("open the selected Contacts folder",()=>session.GetFolderFromID(mailbox.FolderId,mailbox.StoreId));
            items=ReadValue<object>("open the Contacts item collection",()=>folder.Items);
            string Optional(Func<object> read){try{return Text(read());}catch(Exception ex)when(ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){return "unavailable";}}
            ProfileDiagnostics="Outlook version: "+Optional(()=>app.Version)+" | Offline: "+Optional(()=>session.Offline);
            int total=ReadValue<int>("count Contacts items",()=>(int)items.Count);
            for(int i=1;i<=total;i++)
            {
                cancellation.ThrowIfCancellationRequested();progress?.Invoke(i,total);dynamic? item=null;
                try
                {
                    item=ReadValue<object>("open Contacts item "+i+" of "+total,()=>items.Item(i));
                    string kind=ReadValue<string>("read the item type",()=>Text(item.MessageClass));
                    if(kind is not ("IPM.Contact" or "IPM.DistList"))continue;
                    if(kind=="IPM.DistList"&&!ShouldReadGroup(ReadValue<string>("read the contact group name",()=>Text(item.DLName)),Marker(item) is not null))continue;
                    result.Add(ReadItem(item));
                }
                catch(OperationCanceledException){throw;}
                catch(Exception ex){throw new InvalidOperationException("Reading Contacts item "+i+" of "+total+": "+ex.Message,ex);}
                finally{Release(item);}
            }
            return result;
        }
        finally{Release(items);Release(folder);Release(session);Release(app);}
    }
    private static OutlookMarker? Marker(dynamic item)
    {
        dynamic accessor=item.PropertyAccessor;
        try
        {
            string json;
            try{json=Text(accessor.GetProperty(Property));}
            catch(COMException ex)when(ex.ErrorCode==unchecked((int)0x8004010F)){return null;}
            catch(COMException scalarFailure)
            {
                // Some Exchange providers fail the scalar getter for an absent
                // named property. GetProperties exposes the per-property error.
                object? value;
                try
                {
                    var values=(Array)accessor.GetProperties(new object[]{Property});
                    if(values.Length!=1)throw new InvalidOperationException("Outlook returned an invalid tracking-property response.");
                    value=values.GetValue(values.GetLowerBound(0));
                }
                catch(COMException){throw new OutlookReadException("read app tracking information",scalarFailure);}
                int? error=value is ErrorWrapper wrapper?wrapper.ErrorCode:value is int number&&number<0?number:null;
                if(error==unchecked((int)0x8004010F))return null;
                if(error is not null)throw new OutlookReadException("read app tracking information",new COMException("The mailbox could not read the app tracking property.",error.Value));
                json=value as string??throw new OutlookReadException("read app tracking information",scalarFailure);
            }
            var marker=JsonSerializer.Deserialize<OutlookMarker>(json);
            if(marker?.App!=Owner||!Guid.TryParseExact(marker.Id,"N",out _))throw new InvalidOperationException("An Outlook item has invalid app tracking information. Review it before syncing.");
            return marker;
        }
        finally{Release(accessor);}
    }
    private static bool ValidSmtp(string value)=>EmailLists.GeneratedEmailListService.IsValidEmail(value);
    private static string Smtp(dynamic recipient)
    {
        dynamic? entry=null,exchange=null,list=null;
        Exception? failure=null;
        void Remember(Exception ex)
        {
            if(failure is null||failure is COMException missing&&missing.ErrorCode==unchecked((int)0x8004010F))failure=ex;
        }
        string Attempt(Func<object?> read)
        {
            try{string value=Text(read()).Trim().ToLowerInvariant();return ValidSmtp(value)?value:"";}
            catch(Exception ex)when(ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){Remember(ex);return "";}
        }
        string Properties(dynamic owner)
        {
            dynamic? accessor=null;
            try
            {
                accessor=owner.PropertyAccessor;
                foreach(string tag in new[]{"0x39FE001F","0x39FE001E"})
                {string address=Attempt(()=>accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/"+tag));if(address.Length>0)return address;}
            }
            catch(Exception ex)when(ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){Remember(ex);}
            finally{Release(accessor);}
            return "";
        }
        try
        {
            // Read the recipient's native SMTP property before requiring an
            // AddressEntry: legacy one-off list members may not expose one.
            string address=Properties(recipient);if(address.Length>0)return address;
            try{entry=recipient.AddressEntry;}
            catch(Exception ex)when(ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){Remember(ex);}
            if(entry is not null)
            {
                address=Properties(entry);if(address.Length>0)return address;
                string type="";try{type=Text(entry.Type);}catch(COMException ex){Remember(ex);}
                if(type.Equals("EX",StringComparison.OrdinalIgnoreCase))
                {
                    try{exchange=entry.GetExchangeUser();if(exchange is not null){address=Attempt(()=>exchange.PrimarySmtpAddress);if(address.Length>0)return address;}}
                    catch(Exception ex)when(ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){Remember(ex);}
                    try{list=entry.GetExchangeDistributionList();if(list is not null){address=Attempt(()=>list.PrimarySmtpAddress);if(address.Length>0)return address;}}
                    catch(Exception ex)when(ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){Remember(ex);}
                }
                address=Attempt(()=>entry.Address);if(address.Length>0)return address;
            }
            address=Attempt(()=>recipient.Address);if(address.Length>0)return address;
            address=Attempt(()=>recipient.Name);if(address.Length>0)return address;
            throw new OutlookReadException("resolve a group member to an SMTP email address",failure??new InvalidOperationException("No valid SMTP address was available. Review this member in Outlook."));
        }
        finally{Release(list);Release(exchange);Release(entry);}
    }
    private static string ExchangeContact(dynamic item,int slot,string raw,string entryProperty)
    {
        dynamic? session=null,recipient=null,accessor=null;Exception? failure=null;
        try
        {
            // Outlook keeps a slot-specific SMTP counterpart even when an old
            // Exchange directory entry can no longer be resolved.
            accessor=item.PropertyAccessor;
            string smtpLid=(0x8084+slot*0x10).ToString("X4",System.Globalization.CultureInfo.InvariantCulture);
            foreach(string type in new[]{"001F","001E"})
            {
                try
                {
                    string cached=Text(accessor.GetProperty("http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/"+smtpLid+type)).Trim().ToLowerInvariant();
                    if(ValidSmtp(cached))return cached;
                }
                catch(Exception ex)when(ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){failure=ex;}
            }
            session=ReadValue<object>("read the contact MAPI session",()=>item.Session);
            try
            {
                string entryId=Text(accessor.BinaryToString(accessor.GetProperty("http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/"+entryProperty)));
                if(entryId.Length>0){recipient=session.GetRecipientFromID(entryId);string address=Smtp(recipient);if(ValidSmtp(address))return address;}
            }
            catch(Exception ex)when(ex is COMException or OutlookReadException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){failure=ex;}
            finally{Release(recipient);recipient=null;}
            if(raw.Length>0)
            {
                try
                {
                    recipient=session.CreateRecipient(raw);
                    if((bool)recipient.Resolve()){string address=Smtp(recipient);if(ValidSmtp(address))return address;}
                }
                catch(Exception ex)when(ex is COMException or OutlookReadException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){failure=ex;}
            }
            throw new OutlookReadException("resolve Exchange Email"+(slot+1)+" to an SMTP address",failure??new InvalidOperationException("Outlook did not provide an SMTP address."));
        }
        finally{Release(accessor);Release(recipient);Release(session);}
    }
    private const string FileUnderId="http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/80060003";
    private const string FileUnder="http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/8005001F";
    private static (bool Captured,int? Value) ReadFilingFormat(dynamic item)
    {
        dynamic accessor=item.PropertyAccessor;
        try
        {
            object value;
            try{value=accessor.GetProperty(FileUnderId);}
            catch(COMException ex)when(ex.ErrorCode==unchecked((int)0x8004010F)){return (true,null);}
            catch(COMException scalar)
            {
                // Preserve the same Exchange missing-named-property fallback used
                // for app markers. -1/-2/-3 are valid format values, not HRESULTs.
                try{var values=(Array)accessor.GetProperties(new object[]{FileUnderId});if(values.Length!=1)throw new InvalidOperationException("Invalid filing-format response.");value=values.GetValue(values.GetLowerBound(0))!;}
                catch(COMException){throw new OutlookReadException("read the filing-name format",scalar);}
                int? error=value is ErrorWrapper wrapper?wrapper.ErrorCode:value is int number&&number< -3?number:null;
                if(error==unchecked((int)0x8004010F))return (true,null);
                if(error is not null)throw new OutlookReadException("read the filing-name format",new COMException("The filing format could not be read.",error.Value));
            }
            return value is int numberValue?(true,numberValue):value is uint unsigned?(true,unchecked((int)unsigned)):throw new InvalidOperationException("Outlook returned an invalid filing-name format.");
        }
        finally{Release(accessor);}
    }
    private static void SetFilingName(dynamic item,OutlookEntity after)
    {
        string wanted=after.Fields.GetValueOrDefault("FileAs","");
        if(!after.FilingFormatCaptured){if(Text(item.FileAs)!=wanted)item.FileAs=wanted;return;}
        dynamic accessor=item.PropertyAccessor;
        try
        {
            // Set the Unicode value directly and the explicit custom format last,
            // avoiding FileAs's automatic best-match inference from parsed names.
            if(Text(item.FileAs)!=wanted){item.FileAs=wanted;accessor.SetProperty(FileUnder,wanted);}
            (bool Captured,int? Value) current=ReadFilingFormat(item);
            if(current.Value==after.FilingFormat)return;
            if(after.FilingFormat.HasValue)accessor.SetProperty(FileUnderId,after.FilingFormat.Value);
            else try{accessor.DeleteProperty(FileUnderId);}catch(COMException ex)when(ex.ErrorCode==unchecked((int)0x8004010F)){}
        }
        finally{Release(accessor);}
    }
    private static OutlookEntity ReadItem(dynamic item)
    {
        bool group=ReadValue<string>("read the item type",()=>Text(item.MessageClass))=="IPM.DistList";
        var marker=ReadValue<OutlookMarker?>("read app tracking information",()=>Marker(item));
        var fields=new Dictionary<string,string>();var emails=new List<string>{"","",""};var members=new List<GroupMember>();var readIssues=new List<string>();
        if(group)
        {
            fields["FullName"]=ReadValue<string>("read the group name",()=>Text(item.DLName));
            int count=ReadValue<int>("read the group member count",()=>(int)item.MemberCount);
            for(int i=1;i<=count;i++)
            {
                dynamic? recipient=null;
                try
                {
                    recipient=ReadValue<object>("open group member "+i,()=>item.GetMember(i));
                    string email=Smtp(recipient);
                    members.Add(new(email,ReadValue<string>("read group member "+i+" display name",()=>Text(recipient.Name))));
                }
                catch(Exception ex){throw new InvalidOperationException("Group "+fields["FullName"]+", member "+i+" of "+count+": "+ex.Message,ex);}
                finally{Release(recipient);}
            }
        }
        else
        {
            fields["FullName"]=ReadValue<string>("read the contact name",()=>Text(item.FullName));fields["FileAs"]=ReadValue<string>("read the contact filing name",()=>Text(item.FileAs));
            fields["Email1DisplayName"]=ReadValue<string>("read Email1 display name",()=>Text(item.Email1DisplayName));fields["Email2DisplayName"]=ReadValue<string>("read Email2 display name",()=>Text(item.Email2DisplayName));fields["Email3DisplayName"]=ReadValue<string>("read Email3 display name",()=>Text(item.Email3DisplayName));
            emails[0]=ReadValue<string>("read Email1 address",()=>Text(item.Email1Address)).Trim().ToLowerInvariant();emails[1]=ReadValue<string>("read Email2 address",()=>Text(item.Email2Address)).Trim().ToLowerInvariant();emails[2]=ReadValue<string>("read Email3 address",()=>Text(item.Email3Address)).Trim().ToLowerInvariant();
            var types=new[]{ReadValue<string>("read Email1 address type",()=>Text(item.Email1AddressType)),ReadValue<string>("read Email2 address type",()=>Text(item.Email2AddressType)),ReadValue<string>("read Email3 address type",()=>Text(item.Email3AddressType))};
            string[] entryProperties=["80850102","80950102","80A50102"];
            for(int i=0;i<3;i++)if(types[i].Equals("EX",StringComparison.OrdinalIgnoreCase)&&emails[i].Length>0&&!ValidSmtp(emails[i]))
            {
                try{emails[i]=ExchangeContact(item,i,emails[i],entryProperties[i]);}
                catch(OutlookReadException ex){readIssues.Add("Email"+(i+1)+": "+ex.Message+" The original Exchange address was preserved; this contact cannot be changed until it is resolved.");}
            }
        }
        (bool Captured,int? Value) filing=group?(Captured:false,Value:(int?)null):ReadFilingFormat(item);
        return new(){FilingFormatCaptured=filing.Captured,FilingFormat=filing.Value,Id=ReadValue<string>("read the item ID",()=>Text(item.EntryID)),Kind=group?"Group":"Contact",Modified=ReadValue<DateTime>("read the item modification time",()=>(DateTime)item.LastModificationTime).ToUniversalTime().ToString("O"),Fields=fields,Emails=emails,ReadIssues=readIssues,Members=members.OrderBy(m=>m.Email).ThenBy(m=>m.Name).ToList(),ManagedId=marker?.Id??"",Owned=marker?.Owned??false,Links=marker?.Links??[],SourceGroup=marker?.SourceGroup??"",MembersFingerprint=marker?.MembersFingerprint??""};
    }
    private static void SetMarker(dynamic item,OutlookEntity after)
    {
        dynamic accessor=item.PropertyAccessor;
        try
        {
            if(after.Managed)accessor.SetProperty(Property,JsonSerializer.Serialize(new OutlookMarker(Owner,after.ManagedId,after.Owned,after.Links,after.SourceGroup,after.MembersFingerprint)));
            else try{accessor.DeleteProperty(Property);}catch(COMException ex)when(ex.ErrorCode==unchecked((int)0x8004010F)){}
        }
        finally{Release(accessor);}
    }
    public OutlookEntity CaptureBackup(OutlookMailbox mailbox,OutlookEntity before)
    {
        dynamic app=Connect(),session=app.GetNamespace("MAPI"),item=session.GetItemFromID(before.Id,mailbox.StoreId);
        string path=Path.Combine(Path.GetTempPath(),"PropertyManager-Outlook-"+Guid.NewGuid().ToString("N")+".msg");
        try
        {
            if(!before.SameVersion(ReadItem(item)))throw new InvalidOperationException("This Outlook item changed before recovery capture. Preview again.");
            item.SaveAs(path,9); // olMSGUnicode: retain notes, photos and other personal fields.
            if(new FileInfo(path).Length>50*1024*1024)throw new InvalidOperationException("This Outlook item is too large for safe recovery. It has not been deleted.");
            return before with{FullItemBackup=File.ReadAllBytes(path)};
        }
        finally{try{File.Delete(path);}finally{Release(item);Release(session);Release(app);}}
    }
    private static void SetContactLabels(dynamic item,OutlookEntity after)
    {
        // Re-read after a FullName change: Outlook may regenerate dependent labels.
        for(int slot=0;slot<3;slot++)
        {
            if(after.Emails.ElementAtOrDefault(slot) is not {Length:>0})continue;
            string field="Email"+(slot+1)+"DisplayName",wanted=after.Fields.GetValueOrDefault(field,"");
            if(slot==0){if(Text(item.Email1DisplayName)!=wanted)item.Email1DisplayName=wanted;}
            else if(slot==1){if(Text(item.Email2DisplayName)!=wanted)item.Email2DisplayName=wanted;}
            else if(Text(item.Email3DisplayName)!=wanted)item.Email3DisplayName=wanted;
        }
        SetFilingName(item,after);
    }
    private static void SaveContact(dynamic item,OutlookEntity after,bool creating,bool nameChanged)
    {
        if(creating)
        {
            item.Email1Address=after.Emails[0];item.Email2Address=after.Emails[1];item.Email3Address=after.Emails[2];
            item.Categories="PropertyManagerDemo Residents";
        }
        if(creating||nameChanged)item.FullName=after.Name;
        if(creating)
        {
            // Exchange may initialize display labels on the first Save. Persist
            // the tracked baseline first, then set labels on the saved contact.
            // The engine has already journaled this creation before either Save.
            SetMarker(item,after);item.Save();
        }
        SetContactLabels(item,after);
        SetMarker(item,after);item.Save();
    }
    private static void AddNamedGroupMembers(dynamic item,dynamic session,List<GroupMember> members)
    {
        foreach(var member in members)
        {
            dynamic? recipient=null;
            try
            {
                // Open a fully specified one-off, never Resolve against the GAL.
                recipient=session.GetRecipientFromID(OneOffRecipient.EntryId(member));
                if(Smtp(recipient)!=member.Email.ToLowerInvariant()||Text(recipient.Name)!=member.Name)
                    throw new InvalidOperationException("Outlook did not retain the group member's email and unit/name label. The group was not saved.");
                item.AddMember(recipient);
            }
            finally{Release(recipient);}
        }
    }
    private static void UpdateGroupMembers(dynamic item,dynamic session,List<GroupMember> before,List<GroupMember> after)
    {
        var delta=GroupMembershipDelta.Build(before,after);
        if(delta.Remove.Count==0&&delta.Add.Count==0)return;
        // Validate the proposed addresses before changing this unsaved group.
        foreach(var member in delta.Add)OneOffRecipient.EntryId(member);
        int count=ReadValue<int>("check the group member count",()=>(int)item.MemberCount);
        if(count!=before.Count)throw new InvalidOperationException("The group membership changed before editing. Generate a new preview; this group was not saved.");
        var removals=delta.Remove.GroupBy(GroupMembershipDelta.Key).ToDictionary(g=>g.Key,g=>g.Count());
        int remaining=delta.Remove.Count;
        // Reverse indices remain valid as individual native recipients are
        // removed. Do not recreate members whose address AND label are unchanged.
        for(int i=count;i>=1&&remaining>0;i--)
        {
            dynamic? recipient=null;
            try
            {
                recipient=ReadValue<object>("open group member for removal",()=>item.GetMember(i));
                var key=GroupMembershipDelta.Key(new(Smtp(recipient),ReadValue<string>("read group member label",()=>Text(recipient.Name))));
                if(removals.GetValueOrDefault(key)==0)continue;
                item.RemoveMember(recipient);removals[key]--;remaining--;
            }
            finally{Release(recipient);}
        }
        if(remaining>0)throw new InvalidOperationException("A group member changed during editing. Generate a new preview; this group was not saved.");
        AddNamedGroupMembers(item,session,delta.Add);
    }
    public OutlookEntity? Write(OutlookMailbox mailbox,OutlookEntity? before,OutlookEntity? after)
    {
        dynamic app=Connect(),session=app.GetNamespace("MAPI"),folder=session.GetFolderFromID(mailbox.FolderId,mailbox.StoreId);
        dynamic? item=null,items=null;bool saved=false,recreated=false;string? restoreFile=null;
        try
        {
            if(before is not null){item=session.GetItemFromID(before.Id,mailbox.StoreId);
                dynamic parent=item.Parent;try{if(Text(parent.EntryID)!=mailbox.FolderId)throw new InvalidOperationException("The Outlook item moved to another folder. Preview again.");}finally{Release(parent);}
                var fresh=ReadItem(item);
                if(!before.SameVersion(fresh)&&!(after is not null&&before.SameSyncState(fresh)))throw new InvalidOperationException("Outlook changed before saving "+before.Kind+" '"+before.Name+"': "+string.Join(", ",before.ChangedFields(fresh))+". No overwrite was performed.");
                before=fresh;}
            else
            {
                if(after is null)return null;
                if(after.FullItemBackup is not null)
                {
                    restoreFile=Path.Combine(Path.GetTempPath(),"PropertyManager-Restore-"+Guid.NewGuid().ToString("N")+".msg");File.WriteAllBytes(restoreFile,after.FullItemBackup);
                    dynamic original=session.OpenSharedItem(restoreFile);try{item=original.Move(folder);}finally{Release(original);}recreated=true;
                }
                else{items=folder.Items;item=items.Add(after.Kind=="Group"?"IPM.DistList":"IPM.Contact");}
            }
            if(after is null)
            {
                if(before?.Owned!=true)throw new InvalidOperationException("Pre-existing personal contacts and groups cannot be deleted by the app.");
                item.Delete();saved=true;return null;
            }
            folder.ShowAsOutlookAB=true;
            if(recreated){/* Full MSG already restored all fields and membership. */}
            else if(after.Kind=="Group")
            {
                item.DLName=after.Name;
                UpdateGroupMembers(item,session,before?.Members??[],after.Members);
            }
            else
            {
                SaveContact(item,after,before is null,before?.Name!=after.Name);
            }
            if(recreated||after.Kind=="Group"){SetMarker(item,after);item.Save();}
            saved=true;
            dynamic persisted=ReadValue<object>("reopen the saved Outlook item",()=>session.GetItemFromID(Text(item.EntryID),mailbox.StoreId));
            try{return ReadItem(persisted);}finally{Release(persisted);}
        }
        finally
        {
            if(item is not null&&!saved)try{item.Close(1);}catch(COMException){}
            Release(item);Release(items);Release(folder);Release(session);Release(app);
            if(restoreFile is not null)File.Delete(restoreFile);
        }
    }
}
