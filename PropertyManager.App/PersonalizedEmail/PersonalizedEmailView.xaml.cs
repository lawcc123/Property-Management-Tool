using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PropertyManager.App.EmailLists;
using PropertyManager.App.Operations;
using PropertyManager.App.Services;
using Microsoft.Win32;

namespace PropertyManager.App.PersonalizedEmail;
public partial class PersonalizedEmailView : UserControl
{
    private readonly Action accountSettings;
    private readonly ObservableCollection<CampaignRecipient> recipients=[];
    private readonly ObservableCollection<CampaignDocument> documents=[];
    private RecipientFileData? imported;
    private string source="",corporation="";
    private bool loading;
    private TextBox? mergeTarget;
    private string? draftPath;
    private string draftBaseline="";
    private Window OwnerWindow=>Window.GetWindow(this);
    internal PersonalizedEmailView(Action accountSettings,CampaignDeliveryStore? deliveryStore=null,Func<ICampaignTransport>? deliveryTransport=null)
    {
        this.accountSettings=accountSettings;this.deliveryTransport=deliveryTransport??(()=>new CampaignSmtpTransport());this.deliveryStore=deliveryStore??new();InitializeComponent();RecipientGrid.ItemsSource=recipients;DocumentsGrid.ItemsSource=documents;Populate(InitialDraft());RefreshSender();this.deliveryStore.RecoverInterrupted();
        SampleDataButton.Visibility=BuildFeatures.DeveloperToolsEnabled?Visibility.Visible:Visibility.Collapsed;
        documents.CollectionChanged+=(_,e)=>{if(e.NewItems is not null)foreach(CampaignDocument doc in e.NewItems)doc.PropertyChanged+=(_,e)=>{if(e.PropertyName!="SizeText")InvalidatePreview();};RefreshAttachmentSizes();};
        sizeTimer.Tick+=(_,_)=>{if(Steps.SelectedIndex==2)RefreshAttachmentSizes();};Loaded+=(_,_)=>sizeTimer.Start();Unloaded+=(_,_)=>sizeTimer.Stop();UpdateSteps();
        foreach(var box in new[]{SubjectText,BodyText})box.TextChanged+=(_,_)=>InvalidatePreview();SignatureText.TextChanged+=SignatureText_Changed;TestMode.Checked+=(_,_)=>InvalidatePreview();TestMode.Unchecked+=(_,_)=>InvalidatePreview();
    }
    private void Error(Exception ex)=>Warn(ex.Message);
    private void IgnoreWheel(object s,MouseWheelEventArgs e){if(s is ComboBox{IsDropDownOpen:false})e.Handled=true;}
    internal void RefreshSender()
    {
        try{var settings=OperationsSettings.Load();string? username=WindowsCredentialVault.Username(OperationsRunner.EmailTarget);SenderText.Text="From: "+(settings.From.Length>0?settings.From:"not configured")+"\n"+settings.SmtpServer+":"+settings.SmtpPort+" | "+(username is null?"No saved sign-in":"Sign-in: "+username)+"\nShared with Reservations & Reminders. This campaign has its own recipients, subject and message.";}
        catch(Exception ex){SenderText.Text="Sender settings unavailable: "+ex.Message;}
    }
    private void Account_Click(object s,RoutedEventArgs e){accountSettings();RefreshSender();}
    internal void AcceptRecipients(RecipientFileData data,string description,bool showMapping=true)
    {
        if(deliveryBusy)throw new InvalidOperationException("Finish or stop the current email operation before replacing recipients.");
        if(!draftActive)StartDraft(SuggestedDraftName());
        imported=data;source=description;corporation="";loading=true;
        EmailColumn.ItemsSource=data.Columns;UnitColumn.ItemsSource=new[]{"(Not provided)"}.Concat(data.Columns).ToList();NameColumn.ItemsSource=new[]{"(First + last name)"}.Concat(data.Columns).ToList();
        string? Find(params string[] names)=>data.Columns.FirstOrDefault(c=>names.Any(n=>c.Equals(n,StringComparison.OrdinalIgnoreCase)));
        EmailColumn.SelectedItem=Find("Email","Emails","Email Address","EmailAddress");UnitColumn.SelectedItem=Find("Unit","Unit ID","Unit Number")??"(Not provided)";NameColumn.SelectedItem=Find("Names","Name","OwnerName","Full Name")??"(First + last name)";
        MappingPanel.Visibility=showMapping?Visibility.Visible:Visibility.Collapsed;SourceText.Text=source;loading=false;MapRecipients();Steps.SelectedIndex=0;StatusText.Text="Recipient snapshot loaded. Review names and addresses before continuing.";
    }
    private void MapRecipients()
    {
        if(loading||imported is null)return;
        if(EmailColumn.SelectedItem is not string email){StatusText.Text="Choose the email column above.";return;}
        recipients.Clear();foreach(var row in CampaignService.Map(imported,email,UnitColumn.SelectedItem as string??"",NameColumn.SelectedItem as string??"",corporation)){row.PropertyChanged+=(_,_)=>{UpdateCount();InvalidatePreview();};recipients.Add(row);}UpdateFields();UpdateCount();InvalidatePreview();
    }
    private void Mapping_Changed(object s,SelectionChangedEventArgs e){if(IsInitialized)MapRecipients();}
    private void UpdateFields(){FieldCombo.ItemsSource=new[]{"Unit","Names","Corporation","Email"}.Concat(recipients.SelectMany(r=>r.Fields.Keys)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();FieldCombo.SelectedIndex=0;}
    private void UpdateCount()
    {
        RecipientSummary.Text=$"{recipients.Count} rows | {recipients.Count(r=>r.Status=="Ready")} ready | {recipients.Count(r=>r.Status=="Print only")} print only | {recipients.Count(r=>r.Status=="Check email")} need correction";
        RecipientCorporationColumn.Visibility=recipients.Select(r=>r.Corporation).Distinct(StringComparer.OrdinalIgnoreCase).Skip(1).Any()?Visibility.Visible:Visibility.Collapsed;
        EditRecipientDataButton.IsEnabled=recipients.Count>0;
    }
    private void InvalidatePreview(){if(loading||deliveryBusy)return;unlockedStep=Math.Min(unlockedStep,Math.Max(0,Steps.SelectedIndex));InvalidatePrepared();UpdateSteps();PreviewList.ItemsSource=null;ShowMessagePreview(-1);PreviewTo.Text="Build the preview to inspect current messages.";PreviewSubject.Text=PreviewBody.Text="";HtmlMessagePreview.Document=new System.Windows.Documents.FlowDocument();PreviewAttachments.ItemsSource=null;ReviewSummary.Text="Build preview to apply your current selections.";}
    internal void ApplyRecipientEdits(IEnumerable<CampaignRecipient> edited)
    {
        recipients.Clear();foreach(var row in edited){row.PropertyChanged+=(_,_)=>{UpdateCount();InvalidatePreview();};recipients.Add(row);}
        imported=null;MappingPanel.Visibility=Visibility.Collapsed;UpdateFields();UpdateCount();InvalidatePreview();
        StatusText.Text="Recipient snapshot updated. Custom columns are available under Personalization field. Save the draft to keep your changes.";
    }
    private void EditRecipientData_Click(object sender,RoutedEventArgs e)
    {
        RecipientGrid.CommitEdit(DataGridEditingUnit.Cell,true);RecipientGrid.CommitEdit(DataGridEditingUnit.Row,true);
        try{var editor=new RecipientDataEditorWindow(recipients){Owner=OwnerWindow};if(editor.ShowDialog()==true&&editor.Result is not null)ApplyRecipientEdits(editor.Result);}catch(Exception ex){Error(ex);}
    }
    private void Recipient_DoubleClick(object sender,MouseButtonEventArgs e)
    {
        if(e.OriginalSource is not DependencyObject source||ItemsControl.ContainerFromElement(RecipientGrid,source) is not DataGridRow||RecipientGrid.SelectedItem is not CampaignRecipient row)return;
        try{var editor=new RecipientDataEditorWindow(recipients);if(editor.EditRecord(row.RecipientId,OwnerWindow))ApplyRecipientEdits(editor.Session.Recipients());}catch(Exception ex){Error(ex);}
    }
    private void Document_EditEnding(object sender,DataGridCellEditEndingEventArgs e)=>Dispatcher.BeginInvoke(new Action(InvalidatePreview));
    private void Recipient_EditEnding(object s,DataGridCellEditEndingEventArgs e)=>Dispatcher.BeginInvoke(new Action(()=>{UpdateCount();InvalidatePreview();}));
    private bool ReplaceRecipients()=>recipients.Count==0||MessageBox.Show(OwnerWindow,"Replace the current recipient snapshot? Save your draft first if you want to keep it.","Change recipient source",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)==MessageBoxResult.Yes;
    private async void Import_Click(object s,RoutedEventArgs e)
    {
        var picker=new OpenFileDialog{Title="Choose a recipient workbook",Filter="Recipient lists|*.xlsx;*.xlsm;*.xls;*.csv"};if(picker.ShowDialog(OwnerWindow)!=true||!ReplaceRecipients())return;
        try{StatusText.Text="Reading recipient columns...";var data=await RecipientFileLoader.LoadAsync(picker.FileName,OwnerWindow,"Occupants Data");AcceptRecipients(data.Data,Path.GetFileName(picker.FileName)+" / "+data.Sheet);}catch(Exception ex){Error(ex);}
    }
    private void GenerateList_Click(object s,RoutedEventArgs e)
    {
        if(!ReplaceRecipients())return;
        var selector=new GeneralEmailListView();((CheckBox)selector.FindName("DeduplicateCheck")).IsChecked=false;((CheckBox)selector.FindName("GroupCheck")).IsChecked=true;
        var popup=new Window{Owner=OwnerWindow,Title="Choose recipients - Personalized Email",Width=1150,Height=Math.Min(850,SystemParameters.WorkArea.Height-45),MinWidth=900,MinHeight=600,WindowStartupLocation=WindowStartupLocation.CenterOwner,Content=selector,Background=(System.Windows.Media.Brush)FindResource("ThemePage")};
        selector.CampaignRequested+=(data,description)=>{AcceptRecipients(data,description,false);popup.DialogResult=true;};popup.ShowDialog();
    }
#if DEVELOPER_TOOLS
    internal void UseSample()
    {
        AcceptRecipients(new(["Unit","Names","Email","Corporation"],[new(){{"Unit","201"},{"Names","Alex Lee"},{"Email","alex@example.com"},{"Corporation","DEMO 1001"}},new(){{"Unit","518"},{"Names","Alex Lee"},{"Email","alex@example.com"},{"Corporation","DEMO 1001"}},new(){{"Unit","TH 616"},{"Names","Jordan Wong"},{"Email",""},{"Corporation","DEMO 2002"}}]),"Sample data - example addresses, not real residents",false);
        SubjectText.Text="{{Corporation}} - Notice for {{Unit}}";BodyText.Text="Dear {{Names}},\n\nWe are writing about the upcoming work affecting unit {{Unit}}. Please review the enclosed notice.\n\nIf you have any questions, please contact the management office.";StatusText.Text="Sample data loaded. Units 201 and 518 share an address; both notices are retained.";
    }

#endif
    private void Sample_Click(object s,RoutedEventArgs e)
    {
#if DEVELOPER_TOOLS
        if(ReplaceRecipients())UseSample();
#endif
    }
    private void MergeTarget_Focus(object s,RoutedEventArgs e)=>mergeTarget=s as TextBox;
    private void InsertField_Click(object s,RoutedEventArgs e)
    {
        var target=mergeTarget??BodyText;if(FieldCombo.SelectedItem is not string field)return;
        int position=target.SelectionStart;string token="{{"+field+"}}";target.SelectedText=token;target.Focus();target.SelectionStart=position+token.Length;target.SelectionLength=0;
    }
    private void AddDocuments(bool templates)
    {
        var picker=new OpenFileDialog{Title=templates?"Choose one or more Word templates":"Choose common attachments",Filter=templates?"Word templates|*.docx":"All files|*.*",Multiselect=true};if(picker.ShowDialog(OwnerWindow)!=true)return;
        foreach(string path in picker.FileNames.Where(p=>!documents.Any(d=>d.Path.Equals(p,StringComparison.OrdinalIgnoreCase))))documents.Add(new(){Path=path,Personalize=templates,Format=templates?"PDF":Path.GetExtension(path).Equals(".docx",StringComparison.OrdinalIgnoreCase)?"Word":"Original",OutputName=(templates?"{{Unit}} - ":"")+Path.GetFileNameWithoutExtension(path)});InvalidatePreview();
    }
    private void AddTemplates_Click(object s,RoutedEventArgs e)=>AddDocuments(true);
    private void AddFiles_Click(object s,RoutedEventArgs e)=>AddDocuments(false);
    private void RemoveDocument_Click(object s,RoutedEventArgs e){foreach(var doc in DocumentsGrid.SelectedItems.Cast<CampaignDocument>().ToList())documents.Remove(doc);InvalidatePreview();}
    private void CheckTemplate_Click(object s,RoutedEventArgs e)
    {
        try
        {
            DocumentsGrid.CommitEdit();if(DocumentsGrid.SelectedItem is not CampaignDocument{IsWord:true} doc)throw new InvalidOperationException("Select a Word template first.");
            var fields=CampaignService.DocumentFields(doc.Path);var available=FieldCombo.Items.Cast<string>().ToHashSet(StringComparer.OrdinalIgnoreCase);var missing=fields.Where(f=>!available.Contains(f)).ToList();
            TemplateStatus.Text="Fields found: "+(fields.Count==0?"none":string.Join(", ",fields))+". "+(missing.Count==0?"All fields match the recipient columns.":"Choose matching fields for: "+string.Join(", ",missing));
            if(recipients.Count==0)TemplateStatus.Text+=" Load recipients to check their custom columns.";
        }catch(Exception ex){Error(ex);}
    }
    private void TemplatePreview_Click(object s,RoutedEventArgs e)
    {
        try
        {
            if(DocumentsGrid.SelectedItem is not CampaignDocument{IsWord:true} doc)throw new InvalidOperationException("Select a Word template first.");
            RecipientGrid.CommitEdit(DataGridEditingUnit.Cell,true);RecipientGrid.CommitEdit(DataGridEditingUnit.Row,true);
            DocumentsGrid.CommitEdit(DataGridEditingUnit.Cell,true);DocumentsGrid.CommitEdit(DataGridEditingUnit.Row,true);
            new PersonalizedTextPreviewWindow(CampaignService.DocumentText(doc.Path,mergeFieldTokens:doc.Personalize),recipients,doc.Personalize){Owner=OwnerWindow}.ShowDialog();
        }catch(Exception ex){Error(ex);}
    }
    private CampaignDraft ReadDraft()
    {
        RecipientGrid.CommitEdit(DataGridEditingUnit.Cell,true);RecipientGrid.CommitEdit(DataGridEditingUnit.Row,true);DocumentsGrid.CommitEdit(DataGridEditingUnit.Cell,true);DocumentsGrid.CommitEdit(DataGridEditingUnit.Row,true);
        return new(){Title=TitleText.Text,LastExportFolder=exportFolder,CurrentStep=Math.Max(0,Steps.SelectedIndex),UnlockedStep=unlockedStep,Source=source,Subject=SubjectText.Text,Body=BodyText.Text,Signature=SignatureText.Text,SignatureHtml=signatureHtml,TestMode=false,TestInbox=TestInbox.Text,Recipients=recipients.ToList(),Documents=documents.ToList()};
    }
    private void Populate(CampaignDraft draft)
    {
        loading=true;SentCompletionNotice.Visibility=Visibility.Collapsed;exportFolder=draft.LastExportFolder;attachmentProblem=null;TitleText.Text=draft.Title;SubjectText.Text=draft.Subject;BodyText.Text=draft.Body;SignatureText.Text=draft.Signature;signatureHtml=draft.SignatureHtml.Length>0?draft.SignatureHtml:draft.Signature==CampaignSignatureSettings.Default?SignatureHtml.DefaultHtml:SignatureHtml.FromPlain(draft.Signature);RefreshSignatureChoices();UpdateSignaturePreview();TestMode.IsChecked=draft.TestMode;TestInbox.Text=draft.TestInbox;source=draft.Source;SourceText.Text=source.Length>0?source:"No recipient list selected.";
        imported=null;MappingPanel.Visibility=Visibility.Collapsed;recipients.Clear();foreach(var row in draft.Recipients){row.PropertyChanged+=(_,_)=>{UpdateCount();InvalidatePreview();};recipients.Add(row);}documents.Clear();foreach(var doc in draft.Documents)documents.Add(doc);UpdateFields();UpdateCount();loading=false;InvalidatePreview();UpdateSteps();draftBaseline=JsonSerializer.Serialize(ReadDraft());
    }
    private void SaveDraft_Click(object s,RoutedEventArgs e)
    {
        if(TrySaveDraft())MessageBox.Show(OwnerWindow,"Draft '"+TitleText.Text+"' was saved successfully.","Draft saved",MessageBoxButton.OK,MessageBoxImage.Information);
    }
    internal bool HasUnsavedDraft()=>draftActive&&(draftPath is null||JsonSerializer.Serialize(ReadDraft())!=draftBaseline);
    internal void DetachDeletedDraft()
    {
        if(draftPath is not null&&!File.Exists(draftPath))
        {
            draftPath=null;draftBaseline="";
            StatusText.Text="The saved draft was deleted. Your current editor content is kept; Save draft will create a new saved draft.";
        }
    }
    internal void OpenSavedDraft(CampaignDraft draft,string path)
    {
        DeliveryHistoryPanel.Visibility=Visibility.Collapsed;HistoryNavigationButton.Content="View history";
        draftActive=true;unlockedStep=lastStep=0;Populate(draft);draftPath=path;
        RestoreWorkflow(draft);
        draftBaseline=JsonSerializer.Serialize(ReadDraft());
    }
    private void OpenDraft_Click(object s,RoutedEventArgs e)
    {
        var library=new CampaignDraftWindow{Owner=OwnerWindow};
        bool open=library.ShowDialog()==true;
        DetachDeletedDraft();
        if(!open||library.SelectedDraft is null||library.SelectedPath is null)return;
        if(!ConfirmLeave())return;
        OpenSavedDraft(library.SelectedDraft,library.SelectedPath);
    }
    internal void BuildPreview()
    {
        try
        {
            var draft=ReadDraft();var previews=CampaignService.Prepare(draft);InvalidatePrepared();PreviewList.ItemsSource=previews;foreach(var preview in previews.Take(3))PreviewList.SelectedItems.Add(preview);ShowMessagePreview(0);
            ReviewSummary.Text=$"{previews.Count} personalized message previews | {recipients.Count(r=>r.Included&&r.Status=="Print only")} print-only rows | {documents.Count} document(s)/attachment(s) per message. "+(draft.TestMode?"All preview destinations use the test inbox.":"Different units keep their own messages, including shared email addresses.");
            StatusText.Text="Previews ready. Prepare all messages to generate attachments, or send 1-5 selected samples to your test inbox.";
        }catch(Exception ex){InvalidatePreview();Error(ex);ReviewSummary.Text=ex.Message;}
    }
#if DEVELOPER_TOOLS
    internal async Task SimulateSample()
    {
        BuildPreview();var batch=PreviewList.Items.Cast<CampaignPreview>().Take(5).ToList();if(batch.Count==0)return;
        DemoButton.IsEnabled=NextButton.IsEnabled=BackButton.IsEnabled=false;Steps.IsEnabled=false;DemoProgress.Visibility=Visibility.Visible;DemoProgress.Value=0;
        try
        {
            for(int i=0;i<batch.Count;i++){StatusText.Text=$"Simulation {i+1} / {batch.Count}: unit {batch[i].Unit}. No mail server is contacted.";await Task.Delay(120);DemoProgress.Value=(i+1)*100.0/batch.Count;}
            StatusText.Text=$"Simulation complete: {batch.Count} sample message(s) processed locally. Zero emails sent; zero documents generated.";
        }
        finally{DemoProgress.Visibility=Visibility.Collapsed;DemoButton.IsEnabled=NextButton.IsEnabled=Steps.IsEnabled=true;UpdateSteps();}
    }

    private async void Demo_Click(object s,RoutedEventArgs e)=>await SimulateSample();
#else
    private void Demo_Click(object s,RoutedEventArgs e) { }
#endif
    private void BuildPreview_Click(object s,RoutedEventArgs e)=>BuildPreview();
    private void Preview_Changed(object s,SelectionChangedEventArgs e)
    {
        var email=e.AddedItems.OfType<CampaignPreview>().LastOrDefault()??PreviewList.SelectedItem as CampaignPreview;
        if(email is not null)ShowMessagePreview(PreviewList.Items.IndexOf(email));
        else if(PreviewList.Items.Count==0)ShowMessagePreview(-1);
    }
    private void Step_Changed(object s,SelectionChangedEventArgs e)
    {
        if(!ReferenceEquals(e.Source,Steps)||!IsInitialized||updatingSteps)return;
        if(Steps.SelectedIndex>unlockedStep){updatingSteps=true;Steps.SelectedIndex=lastStep;updatingSteps=false;return;}
        lastStep=Math.Max(0,Steps.SelectedIndex);UpdateSteps();if(Steps.SelectedIndex==2)RefreshAttachmentSizes();
    }
    private void UpdateSteps()
    {
        if(BackButton is null||Steps is null||DraftStartPanel is null)return;
        bool history=DeliveryHistoryPanel.Visibility==Visibility.Visible;
        DraftStartPanel.Visibility=!draftActive&&!history?Visibility.Visible:Visibility.Collapsed;
        Steps.Visibility=draftActive&&!history?Visibility.Visible:Visibility.Collapsed;
        WorkflowFooter.Visibility=draftActive||deliveryBusy?Visibility.Visible:Visibility.Collapsed;
        DraftNameText.Text=draftActive?"Draft: "+TitleText.Text:"";DraftNameText.Visibility=draftActive?Visibility.Visible:Visibility.Collapsed;
        SaveDraftButton.IsEnabled=draftActive&&!deliveryBusy;NewDraftButton.IsEnabled=!deliveryBusy;
        for(int i=0;i<Steps.Items.Count;i++){var tab=(TabItem)Steps.Items[i];tab.IsEnabled=draftActive&&!deliveryBusy&&i<=unlockedStep;tab.Opacity=tab.IsEnabled?1:0.42;}
        BackButton.Visibility=history?Visibility.Collapsed:Visibility.Visible;BackButton.IsEnabled=!deliveryBusy&&Steps.SelectedIndex>0;
        NextButton.Visibility=history||Steps.SelectedIndex==4?Visibility.Collapsed:Visibility.Visible;
        NextButton.IsEnabled=draftActive&&!deliveryBusy&&(Steps.SelectedIndex!=2||attachmentProblem is null);
        NextButton.Content=Steps.SelectedIndex switch{0=>"Continue to message",1=>"Continue to documents",2=>"Continue to review & test",_=>"Continue to send live emails"};
        NextButton.ToolTip=attachmentProblem;RefreshExportFolder();
    }
    private void Back_Click(object s,RoutedEventArgs e){if(Steps.SelectedIndex>0)Steps.SelectedIndex--;}
    private async void Next_Click(object s,RoutedEventArgs e)=>await AdvanceStep();
}
