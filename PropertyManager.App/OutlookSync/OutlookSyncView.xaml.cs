using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PropertyManager.App.EmailLists;
using Microsoft.Win32;

namespace PropertyManager.App.OutlookSync;

public partial class OutlookSyncView : UserControl
{
    private readonly IOutlookGateway gateway;
    private readonly RecoveryStore recovery;
    private readonly OutlookStaWorker worker=new();
    private SyncPlan? plan;
    private List<RecoveryRun> historyRuns=[];
    private QuickSyncSettings quickSettings=new();
    private List<FolderShortcut> corporationFolders=[];
    private readonly Action<string> destinationReminder;
    private List<QuickSyncSource>? quickSources;
    private OutlookMailbox? quickDestination;
    private CancellationTokenSource? cancellation;
    private bool busy,loading,compact;
    private string operation="Outlook",lastContext="";
    private string? lastDiagnostic;
    internal bool IsBusy=>busy;
    internal void OpenUpdateSettings()=>SetupExpander.IsExpanded=true;
    private Window OwnerWindow=>Window.GetWindow(this);
    internal OutlookSyncView(IOutlookGateway? gateway=null,RecoveryStore? recovery=null,Action<string>? destinationReminder=null)
    {
        this.gateway=gateway??new OutlookGateway();this.recovery=recovery??new RecoveryStore();InitializeComponent();
        RawErrorButton.Visibility=BuildFeatures.DeveloperToolsEnabled?Visibility.Visible:Visibility.Collapsed;
        this.destinationReminder=destinationReminder??(message=>MessageBox.Show(OwnerWindow,message,"Save Outlook destination",MessageBoxButton.OK,MessageBoxImage.Information));
        foreach(var column in ChangesGrid.Columns.OfType<DataGridTextColumn>())
        {var style=new Style(typeof(TextBlock));style.Setters.Add(new Setter(TextBlock.TextWrappingProperty,TextWrapping.Wrap));style.Setters.Add(new Setter(TextBlock.FontSizeProperty,11.0));column.ElementStyle=style;}
        try{quickSettings=QuickSyncSettings.Load();}catch(Exception ex){StatusText.Text="Could not load update settings: "+ex.Message;}
        if(quickSettings.Destination is { } savedDestination){loading=true;MailboxCombo.ItemsSource=new[]{savedDestination};MailboxCombo.SelectedIndex=0;loading=false;}
        ReloadSavedFolders();UpdateQuickCaption();SetupExpander.IsExpanded=quickSettings.Destination is null;
        try{lastDiagnostic=OutlookDiagnostics.Load();ErrorActions.Visibility=lastDiagnostic is null?Visibility.Collapsed:Visibility.Visible;}catch(Exception ex)when(ex is IOException or UnauthorizedAccessException or System.Security.Cryptography.CryptographicException){}
    }
    internal void ReloadSavedFolders()
    {
        if(busy)return;var previousFolders=corporationFolders.ToList();
        try
        {
            var folders=FolderShortcutStore.Load().Where(f=>System.Text.RegularExpressions.Regex.IsMatch(f.Name,@"DEMO[\s_#-]*(2002|1001)\b",System.Text.RegularExpressions.RegexOptions.IgnoreCase)).ToList();
            corporationFolders=folders;ReloadQuickCorporations(folders);
            if(!previousFolders.SequenceEqual(folders))Invalidate();
        }
        catch(Exception ex){loading=false;StatusText.Text=ex.Message;}
    }
    private void View_SizeChanged(object s,SizeChangedEventArgs e)
    {
        compact=ActualHeight<740;
    }
    internal void Cancel()=>cancellation?.Cancel();
    private void Cancel_Click(object s,RoutedEventArgs e)=>Cancel();
    private void IgnoreWheel(object s,MouseWheelEventArgs e){if(s is ComboBox {IsDropDownOpen:false})e.Handled=true;}
    private void Selection_Changed(object s,SelectionChangedEventArgs e){if(!loading&&IsInitialized)Invalidate();}
    private void Invalidate(){quickSources=null;quickDestination=null;plan=null;ChangesGrid.ItemsSource=null;ApplyButton.IsEnabled=false;SummaryText.Text="Click Update all corporations to load the latest data and update Outlook. Any items needing approval will appear here.";}
    private void SetBusy(bool value)
    {
        busy=value;QuickActions.IsEnabled=SetupExpander.IsEnabled=!value;HistoryNavigationButton.IsEnabled=!value;SettingsPanel.IsEnabled=ActionsPanel.IsEnabled=ChangesGrid.IsEnabled=HistoryPanel.IsEnabled=!value;
        Progress.Visibility=CancelButton.Visibility=value?Visibility.Visible:Visibility.Collapsed;Progress.IsIndeterminate=value;ApplyButton.IsEnabled=!value&&plan is not null&&plan.Changes.Any(r=>r.CanApply&&r.Action!="Unchanged");
        if(value){cancellation=new();ErrorActions.Visibility=Visibility.Collapsed;}else{cancellation?.Dispose();cancellation=null;}
    }
    private IProgress<(string,int,int)> ProgressSink()=>new Progress<(string Message,int Done,int Total)>(p=>{if(!busy)return;StatusText.Text=p.Message+" - "+p.Done+" / "+p.Total;Progress.IsIndeterminate=p.Total==0||p.Message.StartsWith("Checking",StringComparison.Ordinal);Progress.Value=p.Total==0?0:p.Done*100.0/p.Total;});
    private void Error(Exception ex)
    {
        StatusText.Text=ex is OperationCanceledException?"Cancelled. Any saved Outlook changes are listed in Sync history.":"Stopped: "+ex.Message;
        if(ex is not OperationCanceledException)
        {
            lastDiagnostic=OutlookDiagnostics.Record(operation,lastContext+(gateway is OutlookGateway live?" | "+live.ProfileDiagnostics:""),ex);ErrorActions.Visibility=Visibility.Visible;
            StatusText.Text+=" Open Troubleshooting to export a diagnostic report.";
        }
    }
    private void ErrorDetails_Click(object s,RoutedEventArgs e)
    {
#if DEVELOPER_TOOLS
        if(lastDiagnostic is not null)ShowText("Outlook error details",lastDiagnostic);
#endif
    }
    private void ExportError_Click(object s,RoutedEventArgs e)
    {
        if(lastDiagnostic is null)return;
        var picker=new SaveFileDialog{Title="Export Outlook diagnostic report",Filter="JSON diagnostics|*.json",FileName="Outlook-error-details.json",OverwritePrompt=true};
        if(picker.ShowDialog(OwnerWindow)==true)try{File.WriteAllText(picker.FileName,OutlookDiagnostics.ExportReport(lastDiagnostic));}catch(Exception ex){MessageBox.Show(OwnerWindow,ex.Message,"Export diagnostic report");}
    }
    private async void Connect_Click(object s,RoutedEventArgs e)
    {
        if(busy)return;operation="Connect to Classic Outlook";lastContext="Enumerating accessible Contacts folders";try{SetBusy(true);StatusText.Text="Reading accessible Classic Outlook Contacts folders...";var mailboxes=await worker.Run(()=>gateway.Mailboxes());loading=true;MailboxCombo.ItemsSource=mailboxes;MailboxCombo.SelectedItem=mailboxes.FirstOrDefault(b=>b.StoreId==quickSettings.Destination?.StoreId&&b.FolderId==quickSettings.Destination?.FolderId);loading=false;StatusText.Text=MailboxCombo.SelectedItem is null?"Connected. Select your working Contacts folder and Save destination. Archive folders are never selected automatically.":"Connected to your saved Contacts destination.";Invalidate();}catch(Exception ex){Error(ex);}finally{SetBusy(false);}
    }
    private void UpdateQuickCaption()
    {
        DestinationSetupNotice.Visibility=quickSettings.Destination is null?Visibility.Visible:Visibility.Collapsed;
        int count=QuickCorporations.Children.OfType<CheckBox>().Count(c=>c.IsChecked==true);
        QuickDestinationText.Text=(quickSettings.Destination is null?"Destination not set. Open Update settings to choose your working Contacts folder.":"Destination: "+quickSettings.Destination.Label)+" | "+count+" corporation(s) included";
    }
    private void ReloadQuickCorporations(IEnumerable<FolderShortcut> folders)
    {
        QuickCorporations.Children.Clear();
        foreach(string corporation in folders.Select(f=>SyncPlanner.Corporation(f.Name)).Distinct())
        {
            var box=new CheckBox{Content=corporation,Tag=corporation,IsChecked=!quickSettings.ExcludedCorporations.Contains(corporation),Margin=new Thickness(0,0,20,8)};
            box.Checked+=QuickScope_Changed;box.Unchecked+=QuickScope_Changed;QuickCorporations.Children.Add(box);
        }
        UpdateQuickCaption();
    }
    private void QuickScope_Changed(object s,RoutedEventArgs e)
    {
        try{quickSettings.ExcludedCorporations=QuickCorporations.Children.OfType<CheckBox>().Where(c=>c.IsChecked!=true).Select(c=>(string)c.Tag).ToList();quickSettings.Save();UpdateQuickCaption();Invalidate();}
        catch(Exception ex){Error(ex);}
    }
    private void UpdateSettings_Click(object s,RoutedEventArgs e){if(!busy){SetupExpander.IsExpanded=!SetupExpander.IsExpanded;}}
    private void SaveDestination_Click(object s,RoutedEventArgs e)
    {
        if(busy)return;
        try
        {
            var destination=MailboxCombo.SelectedItem as OutlookMailbox??throw new InvalidOperationException("Connect to Outlook and select the Contacts folder to use.");
            if(MessageBox.Show(OwnerWindow,"Use this destination for one-click updates?\n\n"+destination.Label+"\n\nAll selected corporations will update this folder. Previously approved, linked items will be maintained automatically, including contact deletions and group-member removals. Other personal items and group deletions still require review.","Save Outlook destination",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return;
            var saved=new QuickSyncSettings{Destination=destination,ExcludedCorporations=quickSettings.ExcludedCorporations};saved.Save();quickSettings=saved;UpdateQuickCaption();Invalidate();StatusText.Text="Destination saved for this Windows account. Update all corporations is ready.";SetupExpander.IsExpanded=false;
        }catch(Exception ex){Error(ex);}
    }
    private void PresentQuickReviews(QuickSyncResult result)
    {
        var destination=quickDestination??result.Reviews.FirstOrDefault()?.Plan.Mailbox;
        if(destination is null)return;
        plan=QuickSyncService.CombineReviews(destination,result);
        ChangesGrid.ItemsSource=plan.Changes;ApplyPreviewFilter();ApplyButton.IsEnabled=!busy&&plan.Changes.Any(r=>r.CanApply);
        SummaryText.Text=plan.Changes.Count==0?"All included corporations are up to date.":$"All included corporations: {plan.Changes.Count} remaining item(s). Review and approve the changes you want.";
        ReviewInstructions.Text="First-time personal contacts and groups require approval. Unresolved records cannot be applied. No emails are sent.";
    }
    private async void UpdateAll_Click(object s,RoutedEventArgs e)=>await UpdateAll();
    private async Task UpdateAll()
    {
        if(busy)return;
        if(quickSettings.Destination is null)
        {
            SetupExpander.IsExpanded=true;
            StatusText.Text="Save your destination Contacts folder first: connect to Outlook, select the intended folder, then click Save destination.";
            destinationReminder(StatusText.Text);return;
        }
        Invalidate();ShowHistory(false);operation="Update all Outlook corporations";
        try
        {
            var folders=corporationFolders.Where(f=>!quickSettings.ExcludedCorporations.Contains(SyncPlanner.Corporation(f.Name))).ToList();
            if(folders.Count==0)throw new InvalidOperationException("Select at least one corporation in Update settings.");
            if(folders.GroupBy(f=>SyncPlanner.Corporation(f.Name)).Any(g=>g.Count()>1))throw new InvalidOperationException("Duplicate corporation source folders are configured. Resolve them in Manage folders before updating.");
            SetBusy(true);var token=cancellation!.Token;var progress=ProgressSink();StatusText.Text="Connecting to your saved Contacts destination...";
            var boxes=await worker.Run(()=>gateway.Mailboxes());
            if(!boxes.Any(b=>b.StoreId==quickSettings.Destination.StoreId&&b.FolderId==quickSettings.Destination.FolderId))SetupExpander.IsExpanded=true;
            var destination=quickSettings.Resolve(boxes);
            loading=true;MailboxCombo.ItemsSource=boxes;MailboxCombo.SelectedItem=destination;loading=false;lastContext=destination.Label;
            var sources=new List<QuickSyncSource>();var names=GroupNaming.Load();
            foreach(var folder in folders)
            {
                token.ThrowIfCancellationRequested();string corp=SyncPlanner.Corporation(folder.Name);StatusText.Text=corp+" - checking latest Occupant Data...";
                string path=Latest(folder.Path),hash=SyncEngine.Hash(path);var data=await QuickSyncWorkbook.LoadAsync(path);
                if(hash!=SyncEngine.Hash(path))throw new InvalidOperationException(corp+": the workbook changed while loading. Try again.");
                var targets=SyncPlanner.Targets(data,corp);targets=targets with{GroupNames=names.Resolve(corp,targets.Groups.Keys)};
                sources.Add(new(corp,targets,path,hash));
            }
            SourceText.Text=string.Join("\n",sources.Select(source=>source.Corporation+": "+source.Path));
            var result=await worker.Run(()=>new QuickSyncService(gateway,recovery).Run(destination,sources,(m,d,t)=>progress.Report((m,d,t)),token));
            StatusText.Text=$"Finished: {result.Created} created, {result.Updated} updated, {result.Removed} deleted. "+(result.Reviews.Count==0?"Everything is up to date.":"Review the remaining items from all corporations below.");
            if(result.Warnings.Count>0)StatusText.Text+=" "+string.Join(" ",result.Warnings.Take(3));
            quickSources=sources;quickDestination=destination;PresentQuickReviews(result);
            SummaryText.Text+=" "+result.Unchanged+" unchanged contact/group records. Recovery records are available in Sync history.";
            ReviewInstructions.Text="No emails were sent. Existing personal items needing first-time approval were left unchanged.";
            SetupExpander.IsExpanded=false;
        }
        catch(Exception ex){Error(ex);StatusText.Text+=" Any completed changes are recorded in Sync history. Run again after resolving the problem to finish remaining updates.";if(quickSettings.Destination is null)SetupExpander.IsExpanded=true;}
        finally{SetBusy(false);RefreshHistory();}
    }
    internal static string Latest(string folder)=>Directory.EnumerateFiles(folder).Where(p=>!Path.GetFileName(p).StartsWith("~$")&&new[]{".xlsx",".xlsm",".xls"}.Contains(Path.GetExtension(p).ToLowerInvariant())&&new FileInfo(p).Length>0)
        .OrderByDescending(p=>GeneralEmailListView.BackupDateFromFileName(p)??File.GetLastWriteTime(p)).ThenByDescending(p=>p,StringComparer.Ordinal).FirstOrDefault()??throw new InvalidDataException("No Occupant Data workbooks were found in this corporation's folder.");
    private async void FullUndo_Click(object s,RoutedEventArgs e)=>await PreviewFullUndo();
    private async Task PreviewFullUndo()
    {
        if(busy)return;Invalidate();
        try
        {
            var mailbox=MailboxCombo.SelectedItem as OutlookMailbox??throw new InvalidOperationException("Connect to Classic Outlook and select a Contacts folder first.");
            const string scope="*";
            operation="Preview full Outlook undo";lastContext=mailbox.Label+" | "+scope;SetBusy(true);
            var token=cancellation!.Token;var progress=ProgressSink();
            var result=await worker.Run(()=>FullUndoPlanner.Build(recovery,mailbox,scope,gateway.Read(mailbox,(d,t)=>progress.Report(("Checking Outlook contacts",d,t)),token)));
            token.ThrowIfCancellationRequested();plan=result;ChangesGrid.ItemsSource=plan.Changes;ApplyPreviewFilter();
            SummaryText.Text=$"{plan.Changes.Count(r=>r.CanApply&&r.After is null)} remove - {plan.Changes.Count(r=>r.CanApply&&r.After is not null)} restore original - {plan.Changes.Count(r=>!r.CanApply)} need review. Nothing is selected automatically.";
            ReviewInstructions.Text="Full undo restores recorded app changes. Later edits and missing history are left for review. Use View details for the reason.";
            StatusText.Text="Read-only undo preview for "+(scope=="*"?"all corporations":scope)+" in "+mailbox.Label+". "+(recovery.HasLegacyHistory?"Older history may be incomplete; only recoverable changes are offered. A complete pre-app rollback cannot be guaranteed.":"Recovery history is retained for this Windows account. No Outlook changes have been made.");
            ShowHistory(false);if(compact)SetupExpander.IsExpanded=false;
        }
        catch(Exception ex){Error(ex);}finally{SetBusy(false);}
    }
    private void GroupSettings_Click(object s,RoutedEventArgs e)
    {
        if(busy)return;
        try
        {
            if(new GroupSettingsWindow(GroupNaming.Load()){Owner=OwnerWindow}.ShowDialog()==true)
            {Invalidate();StatusText.Text="Group names saved for both corporations. Click Update all corporations to apply updates; first-time changes still need approval.";}
        }catch(Exception ex){Error(ex);}
    }
    private void ApplyPreviewFilter()
    {
        if(ChangesGrid?.ItemsSource is null)return;
        int filter=PreviewFilter.SelectedIndex;
        System.Windows.Data.CollectionViewSource.GetDefaultView(ChangesGrid.ItemsSource).Filter=item=>item is SyncChange row&&(filter<=0||row.ItemKind==(filter==1?"Contact":"Group"));
    }
    private void PreviewFilter_Changed(object s,SelectionChangedEventArgs e){if(IsInitialized)ApplyPreviewFilter();}
    private IEnumerable<SyncChange> DisplayedChanges()=>ChangesGrid.Items.Cast<SyncChange>();
    internal Window CreateEnlargedPreview()
    {
        if(busy||PreviewPanel.Child is null)throw new InvalidOperationException("The preview is unavailable.");
        ChangesGrid.CommitEdit();var content=PreviewPanel.Child;PreviewPanel.Child=null;
        var enlarged=new Window{Owner=OwnerWindow,Title="Outlook Contact Sync - enlarged preview",Width=1200,Height=800,MinWidth=800,MinHeight=500,WindowStartupLocation=WindowStartupLocation.CenterOwner,Background=System.Windows.Media.Brushes.White};
        var root=new DockPanel{Margin=new Thickness(24)};var back=new Button{Content="Back to normal preview",HorizontalAlignment=HorizontalAlignment.Right,Margin=new Thickness(0,0,0,14)};DockPanel.SetDock(back,Dock.Top);root.Children.Add(back);
        var liveStatus=new StackPanel{Margin=new Thickness(0,0,0,12)};DockPanel.SetDock(liveStatus,Dock.Top);
        var status=new TextBlock{TextWrapping=TextWrapping.Wrap};status.SetBinding(TextBlock.TextProperty,new System.Windows.Data.Binding("Text"){Source=StatusText});liveStatus.Children.Add(status);
        var cancel=new Button{Content="Cancel after current item",HorizontalAlignment=HorizontalAlignment.Left,Margin=new Thickness(0,8,0,0)};cancel.SetBinding(VisibilityProperty,new System.Windows.Data.Binding("Visibility"){Source=CancelButton});cancel.Click+=(_,_)=>Cancel();liveStatus.Children.Add(cancel);root.Children.Add(liveStatus);
        root.Children.Add(content);enlarged.Content=root;
        EnlargePreviewButton.Visibility=Visibility.Collapsed;back.Click+=(_,_)=>enlarged.Close();
        enlarged.Closing+=(_,args)=>{if(busy)args.Cancel=true;};
        enlarged.Closed+=(_,_)=>{ChangesGrid.CommitEdit();root.Children.Remove(content);PreviewPanel.Child=content;EnlargePreviewButton.Visibility=Visibility.Visible;};
        return enlarged;
    }
    private void EnlargePreview_Click(object s,RoutedEventArgs e){if(!busy)try{CreateEnlargedPreview().ShowDialog();}catch(Exception ex){Error(ex);}}
    private void ManageFolders_Click(object s,RoutedEventArgs e)
    {if(BuildingLink.CorporationConfiguration.Manage(OwnerWindow)){ReloadSavedFolders();Invalidate();}}
    private void SelectAll_Click(object s,RoutedEventArgs e)
    {
        if(busy||plan is null)return;
        foreach(var row in DisplayedChanges())row.Selected=row.CanApply;
        ChangesGrid.Items.Refresh();
        int selected=plan.Changes.Count(r=>r.Selected),blocked=plan.Changes.Count(r=>!r.CanApply&&r.Action!="Unchanged");
        StatusText.Text=$"{selected} eligible change(s) selected; {blocked} review item(s) remain blocked. Check the selection, then Apply selected changes. Outlook has not been changed.";
    }
    private void SelectRecommended_Click(object s,RoutedEventArgs e){if(busy||plan is null)return;foreach(var row in DisplayedChanges())row.Selected=row.Recommended;ChangesGrid.Items.Refresh();}
    private void Clear_Click(object s,RoutedEventArgs e){if(busy||plan is null)return;foreach(var row in DisplayedChanges())row.Selected=false;ChangesGrid.Items.Refresh();}
    private void EditEnding(object s,DataGridCellEditEndingEventArgs e){}
    private async void Apply_Click(object s,RoutedEventArgs e)
    {
        if(busy||plan is null)return;
        var selected=plan.Changes.Where(r=>r.Selected&&r.CanApply).ToList();
        if(selected.Count==0){StatusText.Text="Select at least one change to apply.";return;}
        int removals=selected.Count(r=>r.After is null),groupRemovals=selected.Where(r=>r.Before?.Kind=="Group").Sum(r=>r.Before!.Members.Count(m=>r.After is null||!r.After.Members.Any(n=>n.Email==m.Email)));
        string question=plan.Mode=="Undo all"
            ? $"Undo app changes in {plan.Mailbox.Label}?\n\nScope: {(plan.Corporation=="*"?"All corporations":plan.Corporation)}\nSelected items: {selected.Count}\nApp-created items to delete: {removals}\nExisting items to restore: {selected.Count-removals}\n\nUnselected and review items will remain unchanged. Older missing recovery history cannot be reconstructed. A recovery record is saved before each change. No emails are sent."
            : $"Update {plan.Mailbox.Label}?\n\nCorporation: {plan.Corporation}\nSelected items: {selected.Count}\nContacts/groups to delete: {removals}\nGroup email memberships to remove: {groupRemovals}\n\nApproving existing personal contacts updates their filing and email display names. Actual names, email addresses, phones and notes are preserved. Recovery records are saved before each change. No emails are sent.";
        if(MessageBox.Show(OwnerWindow,question,plan.Mode=="Undo all"?"Undo app changes":"Apply Outlook changes",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return;
        operation="Apply selected Outlook changes";lastContext=plan.Corporation+" | "+plan.Mailbox.Label;
        var snapshot=plan;var batchSources=quickSources;var batchDestination=quickDestination;bool applied=false;try
        {
            SetBusy(true);var progress=ProgressSink();var engine=new SyncEngine(gateway,recovery);var quickEngine=new QuickSyncService(gateway,recovery);var token=cancellation!.Token;
            string id=await worker.Run(()=>batchSources is not null&&batchDestination is not null
                ?quickEngine.ApplyReviews(snapshot,batchSources,(m,d,t)=>progress.Report((m,d,t)),token)
                :engine.Apply(snapshot,(m,d,t)=>progress.Report((m,d,t)),token));StatusText.Text=snapshot.Mode=="Undo all"?"Selected undo changes completed. Unselected and review items were left unchanged. Sync history contains recovery details.":"Outlook changes completed. Sync history contains the recovery records.";
            applied=true;
            var applyWarnings=batchSources is not null&&batchDestination is not null?quickEngine.LastApplyWarnings:engine.LastApplyWarnings;
            if(applyWarnings.Count>0)StatusText.Text="Finished with "+applyWarnings.Count+" item(s) needing review. "+string.Join(" ",applyWarnings.Take(3))+" Other selected items were processed; no automatic write retries were performed.";
        }
        catch(Exception ex){Error(ex);}finally
        {
            Invalidate();
            if(applied&&batchSources is not null&&batchDestination is not null)
            {
                try
                {
                    var progress=ProgressSink();var token=cancellation!.Token;
                    var remaining=await worker.Run(()=>new QuickSyncService(gateway,recovery).Review(batchDestination,batchSources,(m,d,t)=>progress.Report((m,d,t)),token));
                    quickSources=batchSources;quickDestination=batchDestination;PresentQuickReviews(remaining);
                    if(remaining.Reviews.Count==0){SummaryText.Text="All selected corporations are up to date.";ReviewInstructions.Text="Recovery records are available in Sync history. No emails were sent.";}
                }
                catch(Exception ex){Error(ex);}
            }
            SetBusy(false);RefreshHistory();

        }
    }
    private void Details_Click(object s,RoutedEventArgs e)
    {
        if(ChangesGrid.SelectedItem is not SyncChange row)return;
        if(row.ItemKind=="Group")
        {
            var before=row.Before?.Members??[];var after=row.After?.Members??[];
            string Members(IEnumerable<GroupMember> values)=>string.Join("\n",values.Select(m=>m.Name+" <"+m.Email+">"));
            ShowText("Outlook group details",row.Detail+"\n\nCurrent group: "+row.CurrentName+"\nProposed group: "+row.ProposedName+"\n\nMembers to add:\n"+Members(after.Where(a=>!before.Any(b=>b.Email==a.Email)))+"\n\nMembers to remove:\n"+Members(before.Where(b=>!after.Any(a=>a.Email==b.Email)))+"\n\nProposed members:\n"+Members(after));return;
        }
        ShowText("Outlook change details",row.Detail+"\n\nEmail: "+row.Email+"\nCurrent actual name: "+row.CurrentName+"\nProposed filing name: "+row.ProposedName+"\nCurrent filing name: "+row.CurrentFiling+"\nCurrent filing format: "+(row.Before?.FilingFormatCaptured==true?row.Before.FilingFormat?.ToString()??"Missing":"Not captured in this snapshot")+"\nProposed filing format: "+(row.After?.FilingFormatCaptured==true?row.After.FilingFormat?.ToString()??"Missing":"Unchanged/legacy")+"\nCurrent email display: "+row.CurrentDisplay+"\nProposed email display: "+row.ProposedDisplay+"\n\n"+JsonSerializer.Serialize(new{row.Before,row.After},new JsonSerializerOptions{WriteIndented=true}));
    }
    private void ShowText(string title,string text)
    {new Window{Owner=OwnerWindow,Title=title,Width=850,Height=630,WindowStartupLocation=WindowStartupLocation.CenterOwner,Content=new TextBox{Text=text,IsReadOnly=true,TextWrapping=TextWrapping.Wrap,VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Padding=new Thickness(18)}}.ShowDialog();}
    private void RefreshHistory(){try{historyRuns=recovery.Runs();FilterHistory();}catch(Exception ex){Error(ex);}}
    private void FilterHistory()
    {
        if(HistoryGrid is null||HistorySummary is null)return;
        string? selected=(HistoryGrid.SelectedItem as RecoveryRun)?.Id;
        DateTime? from=HistoryFromDate.SelectedDate,to=HistoryToDate.SelectedDate;
        if(from.HasValue&&to.HasValue&&from.Value.Date>to.Value.Date)
        {
            HistoryGrid.ItemsSource=Array.Empty<RecoveryRun>();HistoryGrid.Height=150;
            HistorySummary.Text="Choose a From date on or before the Through date.";return;
        }
        var matching=SyncHistoryFilter.Matching(historyRuns,from,to);
        var visible=matching.Take(20).ToList();HistoryGrid.ItemsSource=visible;
        HistoryGrid.SelectedItem=visible.FirstOrDefault(r=>r.Id==selected);
        HistoryGrid.Height=Math.Clamp(58+visible.Count*44,150,360);
        HistorySummary.Text=(matching.Count==0?"No sync runs match these dates.":$"Showing {visible.Count} of {matching.Count} matching runs, newest first (up to 20).")+
            " Dates and times use Example City time. Older recovery records remain stored.";
    }
    private void HistoryDate_Changed(object s,SelectionChangedEventArgs e)=>FilterHistory();
    private void ClearHistoryDates_Click(object s,RoutedEventArgs e)
    {HistoryFromDate.SelectedDate=null;HistoryToDate.SelectedDate=null;FilterHistory();}
    private void RefreshHistory_Click(object s,RoutedEventArgs e){if(!busy)RefreshHistory();}
    private void ShowHistory(bool show)
    {
        PreviewPanel.Visibility=show?Visibility.Collapsed:Visibility.Visible;
        HistoryPanel.Visibility=show?Visibility.Visible:Visibility.Collapsed;
        SetupCard.Visibility=show?Visibility.Collapsed:Visibility.Visible;
        HistoryNavigationButton.Content=show?"Back to feature":"View history";
        HistoryNavigationButton.ToolTip=show?"Return to the existing preview and selections without reading Outlook again.":"Open sync history and recovery.";
        if(show)HistoryContextText.Text="Selected Contacts folder: "+((MailboxCombo.SelectedItem as OutlookMailbox)?.Label??"Not selected")+" | Corporation: All";
    }
    private void History_Click(object s,RoutedEventArgs e)
    {
        if(busy)return;
        if(HistoryPanel.Visibility==Visibility.Visible){Back_Click(s,e);return;}
        RefreshHistory();ShowHistory(true);
    }
    private void Back_Click(object s,RoutedEventArgs e){if(!busy)ShowHistory(false);}
    private RecoveryRun? SelectedRun()=>HistoryGrid.SelectedItem as RecoveryRun;
    private void RunDetails_Click(object s,RoutedEventArgs e){if(SelectedRun() is {} run)ShowText("Outlook recovery record",new SyncEngine(gateway,recovery).Report(run));}
    private void Export_Click(object s,RoutedEventArgs e)
    {
        if(SelectedRun() is not {} run)return;var picker=new SaveFileDialog{Title="Export Outlook sync report",Filter="JSON report|*.json",FileName="Outlook-sync-"+run.Id+".json",OverwritePrompt=true};
        if(picker.ShowDialog(OwnerWindow)==true)try{File.WriteAllText(picker.FileName,new SyncEngine(gateway,recovery).Report(run));}catch(Exception ex){Error(ex);}
    }
    private void ExportBackups_Click(object s,RoutedEventArgs e)
    {
        if(busy||SelectedRun() is not {} run)return;
        var backups=recovery.Operations(run.Id).Where(o=>o.Before?.FullItemBackup is not null).ToList();
        if(backups.Count==0){StatusText.Text="This run has no full-item backups. Field changes are available in the recovery details.";return;}
        var picker=new OpenFolderDialog{Title="Choose a folder for exported Outlook recovery items",Multiselect=false};
        if(picker.ShowDialog(OwnerWindow)!=true)return;
        try
        {
            foreach(var op in backups)
            {
                string file=Path.Combine(picker.FolderName,"Outlook-recovery-"+run.Id+"-"+op.Id+"-"+Guid.NewGuid().ToString("N")[..8]+".msg");
                using var output=new FileStream(file,FileMode.CreateNew,FileAccess.Write);output.Write(op.Before!.FullItemBackup!);
            }
            StatusText.Text=$"Exported {backups.Count} full Outlook item backups. These files contain contact details and can be opened in Classic Outlook for manual recovery.";
        }
        catch(Exception ex){Error(ex);}
    }
    private async void Restore_Click(object s,RoutedEventArgs e)
    {
        if(busy||SelectedRun() is not {} run)return;
        if(MessageBox.Show(OwnerWindow,$"Restore the retained Outlook changes from {run.LocalTime}?\n\n{run.Corporation}\n{run.Mailbox}\n\nItems changed since that run are skipped and reported as conflicts. This may update or remove app-created contacts and groups. Pre-existing personal items are not deleted.","Restore Outlook changes",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return;
        try
        {
            operation="Restore Outlook changes";lastContext=run.Corporation+" | "+run.Mailbox;SetBusy(true);var progress=ProgressSink();var token=cancellation!.Token;
            var conflicts=await worker.Run(()=>new SyncEngine(gateway,recovery).Restore(run,(m,d,t)=>progress.Report((m,d,t)),token));
            StatusText.Text=conflicts.Count==0?"Restoration completed.":$"Restored eligible items; {conflicts.Count} conflicts were left unchanged.";
            if(conflicts.Count>0)ShowText("Recovery conflicts",string.Join("\n",conflicts));
        }
        catch(Exception ex){Error(ex);}finally{Invalidate();SetBusy(false);RefreshHistory();}
    }
}
