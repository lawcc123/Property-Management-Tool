using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PropertyManager.App.Services;

namespace PropertyManager.App.Security;

public partial class CredentialManagerWindow : Window
{
    internal bool YardiChanged {get;private set;}
    internal bool BuildingLinkChanged {get;private set;}
    internal CredentialManagerWindow(string? yardiOrigin=null,string? buildingLinkTarget=null,string? reminderEmailTarget=null,int selectedFeature=0)
    {
        InitializeComponent();
        try { string origin=yardiOrigin??YardiCredentialStore.ConfiguredOrigin(); AddService("Yardi P2P",()=>YardiCredentialStore.Username(origin),()=>YardiCredentialStore.Read(origin),
            (username,password)=> {YardiCredentialStore.Save(new(){Origin=origin,Username=username,Password=password});YardiChanged=true;},
            ()=>{YardiCredentialStore.Delete(origin);YardiChanged=true;}); }
        catch(Exception) { ServiceTabs.Items.Add(new TabItem{Header="Yardi P2P",Content=new TextBlock{Text="Yardi sign-in is not configured on this computer."}}); }
        string target=buildingLinkTarget??BuildingLinkCredentialStore.Target;
        AddService("BuildingLink",()=>WindowsCredentialVault.Username(target),()=>WindowsCredentialVault.Read(target),
            (username,password)=>{WindowsCredentialVault.Save(new(){Origin=target,Username=username,Password=password});BuildingLinkChanged=true;},
            ()=>{WindowsCredentialVault.Delete(target);BuildingLinkChanged=true;});
        string emailTarget=reminderEmailTarget??Operations.OperationsRunner.EmailTarget;
        AddService("Email sending (SMTP / app password)",()=>WindowsCredentialVault.Username(emailTarget),()=>WindowsCredentialVault.Read(emailTarget),
            (username,password)=>WindowsCredentialVault.Save(new(){Origin=emailTarget,Username=username,Password=password}),()=>WindowsCredentialVault.Delete(emailTarget));
        AddEmailSettings();ServiceTabs.SelectedIndex=Math.Clamp(selectedFeature,0,2);
    }
    private void AddService(string title,Func<string?> getUsername,Func<LoginCredentials?> read,Action<string,string> save,Action delete)
    {
        var content=new StackPanel();
        content.Children.Add(new TextBlock{Text=title,FontSize=18,FontWeight=FontWeights.SemiBold});
        var status=new TextBlock{Margin=new Thickness(0,7,0,14),TextWrapping=TextWrapping.Wrap,Foreground=new SolidColorBrush(Color.FromRgb(102,121,133))}; content.Children.Add(status);
        var fields=new Grid(); fields.ColumnDefinitions.Add(new(){Width=new GridLength(1,GridUnitType.Star)});fields.ColumnDefinitions.Add(new(){Width=new GridLength(18)});fields.ColumnDefinitions.Add(new(){Width=new GridLength(1,GridUnitType.Star)});
        fields.RowDefinitions.Add(new(){Height=GridLength.Auto});fields.RowDefinitions.Add(new(){Height=GridLength.Auto});
        fields.Children.Add(new TextBlock{Text="Username / email",Margin=new Thickness(0,0,0,7)});
        var passwordLabel=new TextBlock{Text="Password",Margin=new Thickness(0,0,0,7)};Grid.SetColumn(passwordLabel,2);fields.Children.Add(passwordLabel);
        var username=new TextBox{Padding=new Thickness(10)};
        Grid.SetRow(username,1);fields.Children.Add(username);
        var passwordSurface=new Grid();Grid.SetRow(passwordSurface,1);Grid.SetColumn(passwordSurface,2);fields.Children.Add(passwordSurface);
        var password=new PasswordBox{Padding=new Thickness(10)};passwordSurface.Children.Add(password);
        var saved=new TextBlock{Text="\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022  Saved",Margin=new Thickness(12,0,0,0),VerticalAlignment=VerticalAlignment.Center,IsHitTestVisible=false,Foreground=status.Foreground}; passwordSurface.Children.Add(saved);
        System.Windows.Automation.AutomationProperties.SetName(username,title+" username");System.Windows.Automation.AutomationProperties.SetName(password,title+" password");
        string? stored=null;
        void Indicator()=>saved.Visibility=stored is not null && stored==username.Text.Trim() && password.Password.Length==0?Visibility.Visible:Visibility.Collapsed;
        username.TextChanged+=(_,_)=>Indicator();password.PasswordChanged+=(_,_)=>Indicator();
        var actions=new StackPanel{Orientation=Orientation.Horizontal,Margin=new Thickness(0,14,0,0)};
        var saveButton=new Button{Content="Save securely",Background=new SolidColorBrush(Color.FromRgb(20,125,112)),Foreground=Brushes.White,Margin=new Thickness(0,0,8,0),Padding=new Thickness(12,8,12,8)};
        saveButton.SetResourceReference(System.Windows.Controls.Control.BackgroundProperty,"ThemeAccent");saveButton.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty,"ThemeAccentForeground");
        var deleteButton=new Button{Content="Delete saved login",Padding=new Thickness(12,8,12,8)};
        void Refresh()
        {
            stored=getUsername(); username.Text=stored??"";password.Clear();deleteButton.IsEnabled=stored is not null;
            status.Text=stored is null?"No saved login. Enter your account below.":"Using saved login: "+stored; Indicator();
        }
        saveButton.Click+=(_,_)=>
        {
            try {
                string user=username.Text.Trim(),secret=password.Password;
                if(user.Length==0) throw new InvalidOperationException("Enter a username.");
                if(secret.Length==0) {var previous=read();if(previous?.Username!=user) throw new InvalidOperationException("Enter a password for this username.");secret=previous.Password;}
                save(user,secret);Refresh();status.Text+=" Saved securely; the next sign-in will verify it.";
            } catch(Exception ex) {MessageBox.Show(this,ex.Message,"Manage credentials",MessageBoxButton.OK,MessageBoxImage.Warning);}
        };
        deleteButton.Click+=(_,_)=>
        {
            if(MessageBox.Show(this,"Delete the saved "+title+" login for this Windows account?","Delete saved login",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return;
            try{delete();Refresh();}catch(Exception ex){MessageBox.Show(this,ex.Message,"Manage credentials");}
        };
        content.Children.Add(fields);actions.Children.Add(saveButton);actions.Children.Add(deleteButton);content.Children.Add(actions);
        try{Refresh();}catch(Exception){status.Text="Windows could not read this saved login.";saved.Visibility=Visibility.Collapsed;}
        ServiceTabs.Items.Add(new TabItem{Header=title.StartsWith("Email sending")?"Email sending":title,Content=new ScrollViewer{Foreground=new SolidColorBrush(Color.FromRgb(23,43,58)),VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Padding=new Thickness(0,0,12,0),Content=new Border{Style=(Style)FindResource("Card"),Padding=new Thickness(20),Child=content}}});
        Closed+=(_,_)=>password.Clear();
    }
}
