using System.Windows;
namespace PropertyManager.App.PersonalizedEmail;
public partial class PersonalizedEmailView
{
    private int currentPreviewIndex=-1;
    private void ShowMessagePreview(int index)
    {
        if(index<0||index>=PreviewList.Items.Count){currentPreviewIndex=-1;PreviousPreviewButton.IsEnabled=NextPreviewButton.IsEnabled=false;PreviewPosition.Text="No preview records";return;}
        currentPreviewIndex=index;var email=(CampaignPreview)PreviewList.Items[index];
        PreviewTo.Text="To: "+email.To;PreviewSubject.Text=email.Subject;PreviewBody.Text=email.Body;ShowHtmlPreview(email.HtmlBody,email.Body);
        PreviewAttachments.ItemsSource=email.Attachments.Count>0?email.Attachments:["No attachments"];
        PreviousPreviewButton.IsEnabled=index>0;NextPreviewButton.IsEnabled=index<PreviewList.Items.Count-1;
        PreviewPosition.Text=$"Record {index+1} of {PreviewList.Items.Count} - Unit {email.Unit}";
        PreviewList.ScrollIntoView(email);ShowPreparedAttachment();
    }
    private void PreviousPreview_Click(object sender,RoutedEventArgs e){if(currentPreviewIndex>0)ShowMessagePreview(currentPreviewIndex-1);}
    private void NextPreview_Click(object sender,RoutedEventArgs e){if(currentPreviewIndex<PreviewList.Items.Count-1)ShowMessagePreview(currentPreviewIndex+1);}
}
