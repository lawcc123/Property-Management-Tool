using System.IO;
using System.Windows;
using System.Windows.Controls;
namespace PropertyManager.App.PersonalizedEmail;
public partial class PersonalizedEmailView
{
    private string signatureHtml="";
    private CampaignDraft InitialDraft()
    {
        EmailSignature signature=SignatureLibrary.BuiltIn;
        try{signature=SignatureLibraryStore.Load().Default;SignatureHint.Text="Choose a saved signature, or use Manage signatures to create, style and choose your default. Each draft keeps its own signature copy.";}
        catch(Exception ex){SignatureHint.Text="The saved signature library could not be loaded. The styled no-reply default is shown; saved files are kept. "+ex.Message;}
        UseSavedSignatureButton.IsEnabled=File.Exists(CampaignSignatureSettings.PathName);
        return new(){Signature=signature.Text,SignatureHtml=signature.Html};
    }
    private void RefreshSignatureChoices()
    {
        bool was=loading;loading=true;
        try
        {
            var choices=new List<EmailSignature>{new("none","No signature",SignatureHtml.FromPlain(""))};
            try{choices.AddRange(SignatureLibraryStore.Load().All);}catch{choices.Add(SignatureLibrary.BuiltIn);}
            var selected=choices.FirstOrDefault(s=>s.Html==signatureHtml&&s.Text==SignatureText.Text);
            if(selected is null){selected=new("draft","This draft's saved signature",signatureHtml.Length>0?signatureHtml:SignatureHtml.FromPlain(SignatureText.Text));choices.Add(selected);}
            SignatureChoiceCombo.ItemsSource=choices;SignatureChoiceCombo.SelectedItem=selected;
        }
        finally{loading=was;}
    }
    private void UpdateSignaturePreview()=>SignaturePreview.Document=SignatureHtml.Document(signatureHtml.Length>0?signatureHtml:SignatureHtml.FromPlain(SignatureText.Text));
    private void ApplySignature(EmailSignature signature)
    {
        bool was=loading;loading=true;
        try{signatureHtml=signature.Html;SignatureText.Text=signature.Text;RefreshSignatureChoices();UpdateSignaturePreview();}
        finally{loading=was;}
        InvalidatePreview();SignatureHint.Text="Signature applied to this draft. Save draft to keep it; Manage signatures controls the default for new drafts.";
    }
    private void SignatureChoice_Changed(object s,SelectionChangedEventArgs e){if(!loading&&!deliveryBusy&&SignatureChoiceCombo.SelectedItem is EmailSignature signature)ApplySignature(signature);}
    private void ManageSignatures_Click(object s,RoutedEventArgs e)
    {
        if(deliveryBusy)return;
        try{var manager=new SignatureManagerWindow{Owner=OwnerWindow};if(manager.ShowDialog()==true&&manager.Result is {} signature)ApplySignature(signature);}catch(Exception ex){Error(ex);}
    }
    private void SignatureText_Changed(object s,TextChangedEventArgs e)
    {
        if(loading||deliveryBusy)return;
        signatureHtml=SignatureHtml.FromPlain(SignatureText.Text);RefreshSignatureChoices();UpdateSignaturePreview();InvalidatePreview();
    }
    private void ShowHtmlPreview(string html,string plain)=>HtmlMessagePreview.Document=SignatureHtml.Document(html.Length>0?html:SignatureHtml.FromPlain(plain));
    // Compatibility for older saved plain-text settings; the new UI uses the library.
    private void SaveSignature_Click(object sender,RoutedEventArgs e)
    {
        if(deliveryBusy)return;
        try{CampaignSignatureSettings.Save(SignatureText.Text);UseSavedSignatureButton.IsEnabled=true;SignatureHint.Text="Signature saved. Save draft to retain this message too.";}catch(Exception ex){Error(ex);}
    }
    private void UseSavedSignature_Click(object sender,RoutedEventArgs e){if(!deliveryBusy)try{ApplySignature(new("legacy","Previously saved signature",SignatureHtml.FromPlain(CampaignSignatureSettings.Load())));}catch(Exception ex){Error(ex);}}
    private void DefaultSignature_Click(object sender,RoutedEventArgs e){if(!deliveryBusy)ApplySignature(SignatureLibrary.BuiltIn);}
}
