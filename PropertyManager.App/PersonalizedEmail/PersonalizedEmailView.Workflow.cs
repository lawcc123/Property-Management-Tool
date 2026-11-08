using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
namespace PropertyManager.App.PersonalizedEmail;

public partial class PersonalizedEmailView
{
    private bool draftActive,updatingSteps;
    private int unlockedStep,lastStep;
    private string? attachmentProblem;
    private readonly DispatcherTimer sizeTimer=new(){Interval=TimeSpan.FromSeconds(1)};
    internal static string SuggestedDraftName()
    {
        var names=CampaignDraftLibrary.List().Select(e=>e.Title.Trim()).ToHashSet(StringComparer.OrdinalIgnoreCase);
        for(int i=1;;i++){string name="Untitled notice - "+i;if(!names.Contains(name))return name;}
    }
    internal void StartDraft(string name)
    {
        name=name.Trim();if(name.Length==0||name.Length>100||name.Any(char.IsControl))throw new InvalidOperationException("Enter a draft name of 1 to 100 characters.");
        if(CampaignDraftLibrary.List().Any(e=>e.Title.Trim().Equals(name,StringComparison.OrdinalIgnoreCase)))throw new InvalidOperationException("A saved draft already has that name. Choose another name or open the existing draft.");
        SentCompletionNotice.Visibility=Visibility.Collapsed;
        DeliveryHistoryPanel.Visibility=Visibility.Collapsed;HistoryNavigationButton.Content="View history";
        var draft=InitialDraft();draft.Title=name;draftPath=null;draftActive=true;unlockedStep=lastStep=0;
        Populate(draft);Steps.SelectedIndex=0;UpdateSteps();StatusText.Text="Draft created: "+name+". Choose recipients, then use Continue to complete each step.";
    }
    private void StartScreen()
    {
        SentCompletionNotice.Visibility=Visibility.Collapsed;
        DeliveryHistoryPanel.Visibility=Visibility.Collapsed;HistoryNavigationButton.Content="View history";
        draftActive=false;draftPath=null;unlockedStep=lastStep=0;Populate(InitialDraft());Steps.SelectedIndex=0;UpdateSteps();
    }
    internal bool EnsureNamedDraft()
    {
        if(draftActive)return true;
        return RequestNewDraft();
    }
    private bool RequestNewDraft()
    {
        string suggested=SuggestedDraftName();
        var input=new TextBox{Text=suggested,Padding=new Thickness(10),Margin=new Thickness(0,12,0,12),MaxLength=100};
        var error=new TextBlock{TextWrapping=TextWrapping.Wrap,Foreground=System.Windows.Media.Brushes.Firebrick,Margin=new Thickness(0,0,0,12)};
        var create=new Button{Content="Create draft",IsDefault=true,Padding=new Thickness(16,9,16,9),Margin=new Thickness(8,0,0,0)};
        var cancel=new Button{Content="Cancel",IsCancel=true,Padding=new Thickness(16,9,16,9)};
        var panel=new StackPanel{Margin=new Thickness(24)};panel.Children.Add(new TextBlock{Text="Name your new draft",FontSize=22,FontWeight=FontWeights.SemiBold});panel.Children.Add(input);panel.Children.Add(error);
        var buttons=new StackPanel{Orientation=Orientation.Horizontal,HorizontalAlignment=HorizontalAlignment.Right};buttons.Children.Add(cancel);buttons.Children.Add(create);panel.Children.Add(buttons);
        var popup=new Window{Owner=OwnerWindow,Title="New email draft",Width=490,SizeToContent=SizeToContent.Height,ResizeMode=ResizeMode.NoResize,WindowStartupLocation=WindowStartupLocation.CenterOwner,Content=panel,Background=(System.Windows.Media.Brush)FindResource("ThemePage")};
        create.Click+=(_,_)=>{try{StartDraft(input.Text);popup.DialogResult=true;}catch(Exception ex){error.Text=ex.Message;}};
        popup.Loaded+=(_,_)=>{input.Focus();input.SelectAll();};return popup.ShowDialog()==true;
    }
    private void NewDraft_Click(object s,RoutedEventArgs e){if(!deliveryBusy&&ConfirmLeave())RequestNewDraft();}
    internal bool ConfirmLeave(Func<MessageBoxResult>? decision=null)
    {
        if(deliveryBusy){Warn("Finish or stop the current email operation before leaving.");return false;}
        if(!HasUnsavedDraft())return true;
        var answer=decision?.Invoke()??MessageBox.Show(OwnerWindow,"Save changes to '"+TitleText.Text+"' before leaving?\n\nYes: save your draft.\nNo: discard unsaved changes; they will be lost.\nCancel: stay here.","Unsaved email draft",MessageBoxButton.YesNoCancel,MessageBoxImage.Question,MessageBoxResult.Cancel);
        if(answer==MessageBoxResult.Cancel)return false;
        if(answer==MessageBoxResult.Yes)return TrySaveDraft();
        StartScreen();return true;
    }
    private bool TrySaveDraft()
    {
        try
        {
            if(!draftActive)throw new InvalidOperationException("Create or open a draft first.");
            var draft=ReadDraft();
            if(draftPath is null&&CampaignDraftLibrary.List().Any(d=>d.Title.Trim().Equals(draft.Title.Trim(),StringComparison.OrdinalIgnoreCase)))throw new InvalidOperationException("Another saved draft has this name. Open it, or create a new draft with a different name.");
            draftPath=CampaignDraftStore.Save(draft,draftPath);draftBaseline=JsonSerializer.Serialize(draft);StatusText.Text="Draft saved securely for this Windows account. No emails sent.";return true;
        }
        catch(Exception ex){Warn(ex.Message);return false;}
    }
    private void Warn(string message)
    {
        StatusText.Text=message;MessageBox.Show(OwnerWindow,message,"Personalized Email",MessageBoxButton.OK,MessageBoxImage.Warning);
    }
    private void RefreshAttachmentSizes()
    {
        if(loading||deliveryBusy||!draftActive)return;
        bool formatted=false;
        try
        {
            foreach(var doc in documents)doc.RefreshSize();
            var draft=ReadDraft();long raw=documents.Sum(d=>new FileInfo(d.Path).Length),estimate=CampaignMessageSize.Sources(draft);
            AttachmentSizeText.Text=$"Files: {CampaignMessageSize.Display(raw)} | Estimated message: {CampaignMessageSize.Display(estimate)} / 20.00 MB | Remaining: {CampaignMessageSize.Display(Math.Max(0,CampaignMessageSize.Limit-estimate))}. Includes encoding, message and signature. Generated Word/PDF sizes are checked again before sending.";
            formatted=true;attachmentProblem=null;CampaignMessageSize.Validate(estimate);
        }
        catch(Exception ex){attachmentProblem=ex.Message;AttachmentSizeText.Text=formatted?AttachmentSizeText.Text+" "+ex.Message:ex.Message;}
        UpdateSteps();
    }
    private void ValidateStage(int step)
    {
        if(!draftActive)throw new InvalidOperationException("Create or open a named draft first.");
        var draft=ReadDraft();
        if(!draft.Recipients.Any(r=>r.Included&&r.Status=="Ready"))throw new InvalidOperationException("Add and include at least one recipient with a valid email address. Missing-email rows may remain as print-only records.");
        if(draft.Recipients.Any(r=>r.Included&&r.Status=="Check email"))throw new InvalidOperationException("Correct or exclude recipients marked Check email before continuing.");
        if(step==0)return;
        if(string.IsNullOrWhiteSpace(draft.Subject)||string.IsNullOrWhiteSpace(draft.Body))throw new InvalidOperationException("Enter a subject and message before continuing.");
        foreach(var row in draft.Recipients.Where(r=>r.Included&&r.Status=="Ready"))
        {
            string subject=CampaignService.Render(draft.Subject,row.Values());CampaignService.Render(draft.Body+draft.Signature,row.Values());
            if(draft.SignatureHtml.Length>0)SignatureHtml.Parse(SignatureHtml.Render(draft.SignatureHtml,row.Values()));
            if(subject.Contains('\r')||subject.Contains('\n'))throw new InvalidOperationException("The email subject must be on one line.");
        }
        if(step==1)return;
        CampaignMessageSize.Validate(CampaignMessageSize.Sources(draft));CampaignService.Prepare(draft);
    }
    internal async Task AdvanceStep(bool showWarning=true)
    {
        if(deliveryBusy)return;
        try
        {
            int current=Steps.SelectedIndex;ValidateStage(current);
            if(current>=4)return;
            unlockedStep=Math.Max(unlockedStep,current+1);updatingSteps=true;Steps.SelectedIndex=current+1;lastStep=current+1;updatingSteps=false;UpdateSteps();
            if(current==2)BuildPreview();
            if(current==3)await PrepareLiveForSending();
        }
        catch(Exception ex){if(showWarning)Warn(ex.Message);else throw;}
    }
    private void DisabledContinue_Click(object s,MouseButtonEventArgs e)
    {
        var p=e.GetPosition(NextButton);
        if(!NextButton.IsEnabled&&draftActive&&!deliveryBusy&&Steps.SelectedIndex==2&&p.X>=0&&p.Y>=0&&p.X<=NextButton.ActualWidth&&p.Y<=NextButton.ActualHeight)
        {e.Handled=true;Warn(attachmentProblem??"Resolve the attachment problem before continuing.");}
    }
}
