using System.ComponentModel;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace PropertyManager.App.Envelopes;

public partial class AddressBookWindow : Window
{
    internal AddressBook Book {get;}
    private bool loading,dirty,bookChanged,acceptClose;
    private MailingAddress? current;
    private string? logoPng;
    private int kind;
    private List<MailingAddress> Addresses=>kind==0?Book.Contractors:Book.Offices;
    internal AddressBookWindow(AddressBook original)
    {
        Book=JsonSerializer.Deserialize<AddressBook>(JsonSerializer.Serialize(original))!;
        InitializeComponent();loading=true;KindCombo.SelectedIndex=0;loading=false;Reload();Closing+=OnClosing;
    }
    private void IgnoreWheel(object s,MouseWheelEventArgs e){if(s is ComboBox {IsDropDownOpen:false})e.Handled=true;}
    private void Edited(object s,TextChangedEventArgs e){if(!loading)dirty=true;}
    private bool DiscardEditor()=>!dirty||MessageBox.Show(this,"Discard the edits in the address form?","Unsaved address",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)==MessageBoxResult.Yes;
    private void Reload(string? id=null)
    {
        loading=true;AddressList.ItemsSource=Addresses.OrderBy(a=>a.DisplayName,StringComparer.OrdinalIgnoreCase).ToList();
        AddressList.SelectedItem=Addresses.FirstOrDefault(a=>a.Id==id)??AddressList.Items.Cast<MailingAddress>().FirstOrDefault();loading=false;
        ShowAddress(AddressList.SelectedItem as MailingAddress);
    }
    private void ShowAddress(MailingAddress? address)
    {
        loading=true;current=address;CompanyText.Text=address?.Company??"";NicknameText.Text=address?.Nickname??"";StreetText.Text=address?.Street??"";
        CityText.Text=address?.City??"";ProvinceText.Text=address?.Province??"";PostalText.Text=address?.Postal??"";
        logoPng=address?.LogoPng;
        LogoWidthText.Text=(address?.LogoWidthMm??22).ToString(System.Globalization.CultureInfo.InvariantCulture);
        LogoPanel.Visibility=kind==1?Visibility.Visible:Visibility.Collapsed;ShowLogo();
        NicknamePanel.Visibility=kind==1?Visibility.Visible:Visibility.Collapsed;DeleteButton.IsEnabled=address is not null;
        dirty=false;ErrorText.Text="";loading=false;
    }
    private void ShowLogo()
    {
        LogoImage.Source=logoPng is null?null:ReturnAddressLogo.Decode(logoPng);
        LogoStatus.Text=logoPng is null?"No logo selected":"Logo saved with this address. Original image file is no longer needed.";
        RemoveLogoButton.IsEnabled=logoPng is not null;
    }
    internal void SetLogoFromFile(string path)
    {
        logoPng=ReturnAddressLogo.Import(path);ShowLogo();dirty=true;
    }
    private void ChooseLogo_Click(object s,RoutedEventArgs e)
    {
        var picker=new Microsoft.Win32.OpenFileDialog{Title="Choose return-address logo",Filter="Logo images (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp",CheckFileExists=true};
        if(picker.ShowDialog(this)!=true)return;
        try{SetLogoFromFile(picker.FileName);ErrorText.Text="";}catch(Exception ex){ErrorText.Text=ex.Message;}
    }
    private void RemoveLogo_Click(object s,RoutedEventArgs e){logoPng=null;ShowLogo();dirty=true;}
    private void Kind_Changed(object s,SelectionChangedEventArgs e)
    {
        if(loading||!IsInitialized)return;
        if(!DiscardEditor()){loading=true;KindCombo.SelectedIndex=kind;loading=false;return;}
        kind=KindCombo.SelectedIndex;Reload();
    }
    private void Address_Changed(object s,SelectionChangedEventArgs e)
    {
        if(loading||!IsInitialized)return;
        if(!DiscardEditor()){loading=true;AddressList.SelectedItem=current;loading=false;return;}
        ShowAddress(AddressList.SelectedItem as MailingAddress);
    }
    private void Add_Click(object s,RoutedEventArgs e)
    {if(!DiscardEditor())return;loading=true;AddressList.SelectedItem=null;loading=false;ShowAddress(null);CompanyText.Focus();}
    private bool SaveAddress()
    {
        try
        {
            if(!double.TryParse(LogoWidthText.Text,System.Globalization.NumberStyles.Float,System.Globalization.CultureInfo.CurrentCulture,out double logoWidth)
                &&!double.TryParse(LogoWidthText.Text,System.Globalization.NumberStyles.Float,System.Globalization.CultureInfo.InvariantCulture,out logoWidth))
                throw new InvalidOperationException("Enter the logo width in millimetres.");
            var address=new MailingAddress{Id=current?.Id??Guid.NewGuid().ToString("N"),Company=CompanyText.Text.Trim(),Nickname=NicknameText.Text.Trim(),Street=StreetText.Text.Trim(),City=CityText.Text.Trim(),Province=ProvinceText.Text.Trim(),Postal=PostalText.Text.Trim(),LogoPng=kind==1?logoPng:current?.LogoPng,LogoWidthMm=logoWidth};
            address.Validate();int index=current is null?-1:Addresses.FindIndex(a=>a.Id==current.Id);
            if(index<0)Addresses.Add(address);else Addresses[index]=address;bookChanged=true;dirty=false;Reload(address.Id);return true;
        }catch(Exception ex){ErrorText.Text=ex.Message;return false;}
    }
    private void SaveAddress_Click(object s,RoutedEventArgs e)=>SaveAddress();
    private void Delete_Click(object s,RoutedEventArgs e)
    {
        if(current is null||MessageBox.Show(this,"Delete this address from the address book? Existing print history is retained.","Delete address",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return;
        Addresses.RemoveAll(a=>a.Id==current.Id);bookChanged=true;dirty=false;Reload();
    }
    private void SaveBook_Click(object s,RoutedEventArgs e)
    {if(dirty&&!SaveAddress())return;Book.Validate();acceptClose=true;DialogResult=true;}
    private void Cancel_Click(object s,RoutedEventArgs e)=>Close();
    private void ImportCsv_Click(object s,RoutedEventArgs e)
    {
        if(!DiscardEditor())return;
        var picker=new Microsoft.Win32.OpenFileDialog{Title="Import mailing addresses",Filter="CSV (*.csv)|*.csv",CheckFileExists=true};
        if(picker.ShowDialog(this)!=true)return;
        try{var incoming=AddressCsv.Read(picker.FileName);int added=AddressCsv.Merge(Addresses,incoming);bookChanged|=added>0;dirty=false;Reload();ErrorText.Text=$"{added} addresses added. Existing company/street duplicates were skipped. Save the address book to keep the changes.";}
        catch(Exception ex){ErrorText.Text=ex.Message;}
    }
    private void ExportCsv_Click(object s,RoutedEventArgs e)
    {
        var picker=new Microsoft.Win32.SaveFileDialog{Title="Export mailing addresses",Filter="CSV (*.csv)|*.csv",DefaultExt=".csv",FileName=kind==0?"Contractors.csv":"Return-addresses.csv",OverwritePrompt=true};
        if(picker.ShowDialog(this)!=true)return;
        try{System.IO.File.WriteAllText(picker.FileName,AddressCsv.Write(Addresses),new System.Text.UTF8Encoding(true));}
        catch(Exception ex){ErrorText.Text=ex.Message;}
    }
    private void OnClosing(object? sender,CancelEventArgs e)
    {if(!acceptClose&&(dirty||bookChanged)&&MessageBox.Show(this,"Discard the unsaved address-book changes?","Unsaved address book",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)e.Cancel=true;}
}
