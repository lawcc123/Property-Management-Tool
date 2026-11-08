using System.Windows;
using System.Windows.Media;
using PropertyManager.App.EmailLists;
using PropertyManager.App.BuildingLink;
using PropertyManager.App.Security;
using PropertyManager.App.Envelopes;

namespace PropertyManager.App;

public partial class MainWindow
{
    private PersonalizedEmail.PersonalizedEmailView? personalizedEmailView;
    private void EnsurePersonalizedEmail(){if(personalizedEmailView is null){personalizedEmailView=new(()=>OpenCredentialManager(2));PersonalizedEmailHost.Content=personalizedEmailView;}}
    private void PersonalizedEmailNavigation_Click(object sender,RoutedEventArgs e)=>ShowPage(7);
    private void UseCampaignList(RecipientFileData data,string description){EnsurePersonalizedEmail();ShowPage(7);if(activeWorkspace==7&&personalizedEmailView!.EnsureNamedDraft())personalizedEmailView.AcceptRecipients(data,description,false);}
    private Operations.OperationsView? operationsView;
    private void OperationsNavigation_Click(object sender,RoutedEventArgs e)=>ShowPage(6);
    private OutlookSync.OutlookSyncView? outlookSyncView;
    private void OutlookSyncNavigation_Click(object sender,RoutedEventArgs e)=>ShowPage(5);
    private GeneralEmailListView? emailListView;
    private GeneralEmailListView? ownerMailingView;
    private BuildingLinkView? buildingLinkView;
    private EnvelopeView? envelopeView;
    private void OwnerMailingNavigation_Click(object sender,RoutedEventArgs e)=>ShowPage(4);
    private void RegistryNavigation_Click(object sender,RoutedEventArgs e)=>ShowPage(0);
    private void EmailListNavigation_Click(object sender,RoutedEventArgs e)=>ShowPage(1);
    private void BuildingLinkNavigation_Click(object sender,RoutedEventArgs e)=>ShowPage(2);
    private void EnvelopeNavigation_Click(object sender,RoutedEventArgs e)=>ShowPage(3);
    internal bool CorporationSettingsBusy=>emailListView?.IsBusy==true||ownerMailingView?.IsBusy==true||running||buildingLinkView?.IsRunning==true||buildingLinkView?.ScheduleBusy==true||outlookSyncView?.IsBusy==true||operationsView?.IsBusy==true;
    private void ShowPage(int page)
    {
        if(page!=7&&activeWorkspace==7&&personalizedEmailView is not null&&!personalizedEmailView.ConfirmLeave())return;
        if(page==7){EnsurePersonalizedEmail();personalizedEmailView!.RefreshSender();}
        PersonalizedEmailHost.Visibility=page==7?Visibility.Visible:Visibility.Collapsed;PersonalizedEmailNavigation.Background=PageBrush(page==7);
        if(page==6&&operationsView is null){try{operationsView=new(()=>OpenCredentialManager(2));OperationsHost.Content=operationsView;}catch(Exception ex){MessageBox.Show(this,ex.Message,"Reservations & Reminders");return;}}
        OperationsHost.Visibility=page==6?Visibility.Visible:Visibility.Collapsed;OperationsNavigation.Background=PageBrush(page==6);
        if(page==5&&outlookSyncView is null){try{outlookSyncView=new();OutlookSyncHost.Content=outlookSyncView;}catch(Exception ex){MessageBox.Show(this,ex.Message,"Outlook Contact Sync");return;}}
        if(page==5)outlookSyncView?.ReloadSavedFolders();
        if(page==1 && emailListView is null){emailListView=new();emailListView.CampaignRequested+=UseCampaignList;EmailListHost.Content=emailListView;}
        if(page==2 && buildingLinkView is null){buildingLinkView=new(()=>OpenCredentialManager(1));BuildingLinkHost.Content=buildingLinkView;}
        if(page==2)buildingLinkView?.ReloadSharedFolders();
        if(page==3 && envelopeView is null){envelopeView=new();EnvelopeHost.Content=envelopeView;}
        if(page==4 && ownerMailingView is null){ownerMailingView=new(true);OwnerMailingHost.Content=ownerMailingView;}
        if(page==1)emailListView?.ReloadSavedFolders();
        if(page==4)ownerMailingView?.ReloadSavedFolders();
        OutlookSyncHost.Visibility=page==5?Visibility.Visible:Visibility.Collapsed;
        OutlookSyncNavigation.Background=PageBrush(page==5);
        WorkspaceScroll.Visibility=page==0?Visibility.Visible:Visibility.Collapsed;
        EmailListHost.Visibility=page==1?Visibility.Visible:Visibility.Collapsed;
        BuildingLinkHost.Visibility=page==2?Visibility.Visible:Visibility.Collapsed;
        EnvelopeHost.Visibility=page==3?Visibility.Visible:Visibility.Collapsed;
        OwnerMailingHost.Visibility=page==4?Visibility.Visible:Visibility.Collapsed;
        WorkspaceBreadcrumb.Text="Workspace  /  "+(page==0?"Registry Sync":page==1?"Generate Email List":page==2?"BuildingLink Data Download":page==3?"Printing contractor's address":page==4?"Owner Mailing List":page==5?"Outlook Contact Sync":page==7?"Personalized Email":"Reservations & Reminders");
        activeWorkspace=page;RegistryNavigation.Background=PageBrush(page==0);EmailListNavigation.Background=PageBrush(page==1);BuildingLinkNavigation.Background=PageBrush(page==2);EnvelopeNavigation.Background=PageBrush(page==3);OwnerMailingNavigation.Background=PageBrush(page==4);
    }
    private static SolidColorBrush PageBrush(bool selected)=>selected?(SolidColorBrush)Application.Current.FindResource("ThemeSidebarSelected"):new(Colors.Transparent);
    private void ManageCredentials_Click(object sender,RoutedEventArgs e)=>OpenCredentialManager();
    private void OpenCredentialManager(int selectedFeature=0,Window? owner=null)
    {
        if(operationsView?.IsBusy==true || running || buildingLinkView?.IsRunning==true || buildingLinkView?.ScheduleBusy==true){MessageBox.Show(this,"Finish the current sync or download before changing saved credentials.","Manage credentials");return;}
        var manager=new CredentialManagerWindow(string.IsNullOrEmpty(loginOrigin)?null:loginOrigin,selectedFeature:selectedFeature){Owner=owner??this};manager.ShowDialog();
        if(manager.YardiChanged)
        {
            savedLoginUsername=null;LoginUsername.Clear();LoginPassword.Clear();ForgetLoginButton.IsEnabled=false;
            UseAutofill.IsChecked=RememberLogin.IsChecked=false;InitializeLogin(loginOrigin);UpdateSavedPasswordIndicator();
        }
        buildingLinkView?.RefreshCredentials();
        personalizedEmailView?.RefreshSender();
        operationsView?.ReloadEmailSettings();
        operationsView?.RefreshCredentials();
    }
}
