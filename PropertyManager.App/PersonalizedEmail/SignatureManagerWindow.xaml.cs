using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using Microsoft.Win32;
namespace PropertyManager.App.PersonalizedEmail;
public partial class SignatureManagerWindow:Window
{
    readonly SignatureLibrary library;
    EmailSignature? selected;
    Image? selectedPicture;
    bool loading=true,dirty,saved;
    internal EmailSignature? Result{get;private set;}
    internal SignatureManagerWindow()
    {
        InitializeComponent();library=SignatureLibraryStore.Load();Render(library.DefaultId);loading=false;
        Closing+=(_,e)=>{if(dirty&&!saved&&MessageBox.Show(this,"Discard unsaved signature changes?","Manage signatures",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)e.Cancel=true;};
    }
    void Error(Exception ex)=>MessageBox.Show(this,ex.Message,"Manage signatures",MessageBoxButton.OK,MessageBoxImage.Warning);
    void Render(string id){bool was=loading;loading=true;SignatureList.ItemsSource=library.All;SignatureList.SelectedItem=library.All.First(s=>s.Id==id);DefaultHint.Text="Default: "+library.Default.Name;loading=was;LoadSelected();}
    void LoadSelected()
    {
        if(SignatureList.SelectedItem is not EmailSignature signature)return;
        loading=true;selected=signature;NameText.Text=signature.Name;Editor.Document=SignatureHtml.Document(signature.Html);bool builtin=signature.Id=="no-reply";Editor.IsReadOnly=builtin;NameText.IsEnabled=FormattingPanel.IsEnabled=DeleteButton.IsEnabled=!builtin;selectedPicture=null;
        EditorHint.Text=builtin?"Built-in no-reply signature. Choose New signature / copy to customize a copy.":"Select text to style it. Click a picture to resize it. Fields such as {{Unit}} can be typed into the signature.";
        BindPictures();loading=false;
    }
    void Commit()
    {
        if(selected is null||selected.Id=="no-reply")return;
        string name=NameText.Text.Trim();if(name.Length==0||name.Length>80)throw new InvalidDataException("Enter a signature name.");
        if(library.All.Any(s=>s.Id!=selected.Id&&s.Name.Equals(name,StringComparison.OrdinalIgnoreCase)))throw new InvalidDataException("Choose a unique signature name.");
        var replacement=new EmailSignature(selected.Id,name,SignatureHtml.Html(Editor.Document));int index=library.Items.FindIndex(s=>s.Id==selected.Id);library.Items[index]=replacement;selected=replacement;
    }
    void Selection_Changed(object s,SelectionChangedEventArgs e)
    {
        if(loading||Editor is null)return;
        var next=SignatureList.SelectedItem as EmailSignature;
        try{Commit();if(next is not null)Render(next.Id);}
        catch(Exception ex){loading=true;SignatureList.SelectedItem=SignatureList.Items.Cast<EmailSignature>().FirstOrDefault(i=>i.Id==selected?.Id);loading=false;Error(ex);}
    }
    void New_Click(object s,RoutedEventArgs e)
    {
        try{Commit();int n=1;while(library.All.Any(s=>s.Name.Equals("Signature "+n,StringComparison.OrdinalIgnoreCase)))n++;var signature=new EmailSignature(Guid.NewGuid().ToString("N"),"Signature "+n,selected?.Html??SignatureHtml.DefaultHtml);library.Items.Add(signature);dirty=true;Render(signature.Id);NameText.Focus();NameText.SelectAll();}catch(Exception ex){Error(ex);}
    }
    void Delete_Click(object s,RoutedEventArgs e)
    {
        if(selected is null||selected.Id=="no-reply")return;
        if(MessageBox.Show(this,"Delete '"+selected.Name+"'? Saved email drafts keep their own signature copy.","Delete signature",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return;
        library.Items.RemoveAll(i=>i.Id==selected.Id);if(library.DefaultId==selected.Id)library.DefaultId="no-reply";dirty=true;selected=null;Render(library.DefaultId);
    }
    void Default_Click(object s,RoutedEventArgs e){try{Commit();if(selected is null)return;library.DefaultId=selected.Id;dirty=true;Render(selected.Id);}catch(Exception ex){Error(ex);}}
    void Save_Click(object s,RoutedEventArgs e){try{Commit();SignatureLibraryStore.Save(library);Result=selected??library.Default;saved=true;DialogResult=true;}catch(Exception ex){Error(ex);}}
    void Name_Changed(object s,TextChangedEventArgs e){if(!loading)dirty=true;}
    void Editor_Changed(object s,TextChangedEventArgs e){if(!loading){dirty=true;BindPictures();}}
    void Command(RoutedUICommand command){if(loading||Editor is null)return;command.Execute(null,Editor);Editor.Focus();dirty=true;}
    void Bold_Click(object s,RoutedEventArgs e)=>Command(EditingCommands.ToggleBold);
    void Italic_Click(object s,RoutedEventArgs e)=>Command(EditingCommands.ToggleItalic);
    void Underline_Click(object s,RoutedEventArgs e)=>Command(EditingCommands.ToggleUnderline);
    void Left_Click(object s,RoutedEventArgs e)=>Command(EditingCommands.AlignLeft);
    void Centre_Click(object s,RoutedEventArgs e)=>Command(EditingCommands.AlignCenter);
    void Right_Click(object s,RoutedEventArgs e)=>Command(EditingCommands.AlignRight);
    void Font_Changed(object s,SelectionChangedEventArgs e){if(!loading&&Editor is not null&&FontCombo.SelectedItem is ComboBoxItem item){Editor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty,new System.Windows.Media.FontFamily(item.Content.ToString()!));dirty=true;}}
    void Size_Changed(object s,SelectionChangedEventArgs e){if(!loading&&Editor is not null&&FontSizeCombo.SelectedItem is ComboBoxItem item){Editor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty,double.Parse(item.Content.ToString()!)*96/72);dirty=true;}}
    void Colour_Click(object s,RoutedEventArgs e){using var dialog=new System.Windows.Forms.ColorDialog();if(dialog.ShowDialog()!=System.Windows.Forms.DialogResult.OK)return;var c=dialog.Color;Editor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty,new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(c.R,c.G,c.B)));dirty=true;}
    double WidthValue(){if(!double.TryParse(PictureWidth.Text,out double width)||width<16||width>600)throw new InvalidDataException("Use a picture width from 16 to 600 pixels.");return width;}
    void BindPictures()
    {
        if(Editor is null)return;
        IEnumerable<Inline> All(IEnumerable<Inline> items){foreach(var inline in items){yield return inline;if(inline is Span span)foreach(var child in All(span.Inlines))yield return child;}}
        foreach(var p in Editor.Document.Blocks.OfType<Paragraph>())foreach(var container in All(p.Inlines).OfType<InlineUIContainer>())if(container.Child is Image image)
        {image.MouseLeftButtonDown-=PictureSelected;image.MouseLeftButtonDown+=PictureSelected;}
    }
    void PictureSelected(object s,MouseButtonEventArgs e){selectedPicture=(Image)s;PictureWidth.Text=selectedPicture.Width.ToString("0");Status.Text="Picture selected. Set its width and choose Resize selected picture.";}
    void Picture_Click(object s,RoutedEventArgs e)
    {
        try
        {
            var picker=new OpenFileDialog{Title="Choose a signature picture",Filter="Pictures|*.png;*.jpg;*.jpeg"};if(picker.ShowDialog(this)!=true)return;
            string source=SignatureHtml.ImportPicture(picker.FileName);var image=new Image{Source=SignatureHtml.DecodeImage(source),Width=WidthValue(),Stretch=System.Windows.Media.Stretch.Uniform,Tag=source,ToolTip="Signature picture"};
            Editor.Selection.Text="";new InlineUIContainer(image,Editor.CaretPosition);selectedPicture=image;dirty=true;BindPictures();
        }
        catch(Exception ex){Error(ex);}
    }
    void Resize_Click(object s,RoutedEventArgs e){try{if(selectedPicture is null)throw new InvalidOperationException("Click a picture in the editor first.");selectedPicture.Width=WidthValue();dirty=true;}catch(Exception ex){Error(ex);}}
    void Link_Click(object s,RoutedEventArgs e)
    {
        var input=new TextBox{Text="https://",Margin=new Thickness(16),Padding=new Thickness(8)};var add=new Button{Content="Insert link",Margin=new Thickness(16)};var panel=new StackPanel();panel.Children.Add(input);panel.Children.Add(add);
        var popup=new Window{Owner=this,Title="Signature link",Width=440,SizeToContent=SizeToContent.Height,WindowStartupLocation=WindowStartupLocation.CenterOwner,Content=panel};
        add.Click+=(_,_)=>{if(!SignatureHtml.SafeLink(input.Text)){MessageBox.Show(popup,"Enter an https, http or mailto address.");return;}string text=Editor.Selection.Text;Editor.Selection.Text="";var link=new Hyperlink(new Run(text.Length==0?input.Text:text),Editor.CaretPosition){NavigateUri=new Uri(input.Text)};dirty=true;popup.DialogResult=true;};popup.ShowDialog();
    }
}
