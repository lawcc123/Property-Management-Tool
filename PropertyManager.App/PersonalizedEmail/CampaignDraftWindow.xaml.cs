using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
namespace PropertyManager.App.PersonalizedEmail;
public partial class CampaignDraftWindow:Window
{
    private List<CampaignDraftEntry> entries=[];
    private bool busy;
    internal Task LoadTask {get;private set;}=Task.CompletedTask;
    internal CampaignDraft? SelectedDraft {get;private set;}
    internal string? SelectedPath {get;private set;}
    public CampaignDraftWindow()
    {
        InitializeComponent();Loaded+=(_,_)=>LoadTask=RefreshAsync();
    }
    internal async Task RefreshAsync()
    {
        if(busy)return;
        busy=true;UpdateActions();RefreshButton.IsEnabled=false;
        try{entries=await Task.Run(CampaignDraftLibrary.List);Filter();}
        catch(Exception ex){StatusText.Text="Could not read saved drafts: "+ex.Message;}
        finally{busy=false;RefreshButton.IsEnabled=true;UpdateActions();}
    }
    private void Filter()
    {
        if(DraftGrid is null)return;
        string query=SearchText.Text.Trim();
        var visible=entries.Where(e=>e.Title.Contains(query,StringComparison.OrdinalIgnoreCase)||e.Subject.Contains(query,StringComparison.OrdinalIgnoreCase)).ToList();
        DraftGrid.ItemsSource=visible;CountText.Text=$"{visible.Count} of {entries.Count} saved drafts";
        EmptyText.Visibility=visible.Count==0?Visibility.Visible:Visibility.Collapsed;
        EmptyText.Text=entries.Count==0?"No saved drafts yet. Use Save draft to keep your work.":"No drafts match your search.";
        UpdateActions();
    }
    private void UpdateActions()
    {
        if(DraftGrid is null)return;
        var selected=DraftGrid.SelectedItems.Cast<CampaignDraftEntry>().ToList();
        OpenButton.IsEnabled=!busy&&selected.Count==1&&selected[0].CanOpen;
        DeleteButton.IsEnabled=!busy&&selected.Count>0;
        DeleteButton.Content=selected.Count>0?$"Delete selected ({selected.Count})":"Delete selected";
    }
    private void Search_Changed(object sender,TextChangedEventArgs e)=>Filter();
    private void Selection_Changed(object sender,SelectionChangedEventArgs e)=>UpdateActions();
    private async void Refresh_Click(object sender,RoutedEventArgs e)=>await RefreshAsync();
    internal bool TryOpenSelected()
    {
        if(busy||DraftGrid.SelectedItems.Count!=1||DraftGrid.SelectedItem is not CampaignDraftEntry {CanOpen:true} entry)return false;
        try{SelectedDraft=CampaignDraftLibrary.Load(entry);SelectedPath=entry.Path;return true;}
        catch(Exception ex){StatusText.Text="Could not open this draft: "+ex.Message;return false;}
    }
    private void Open_Click(object sender,RoutedEventArgs e){if(TryOpenSelected())DialogResult=true;}
    private void Draft_DoubleClick(object sender,MouseButtonEventArgs e)
    {
        if(e.OriginalSource is DependencyObject source&&ItemsControl.ContainerFromElement(DraftGrid,source) is DataGridRow&&OpenButton.IsEnabled)Open_Click(sender,e);
    }
    internal async Task<int> DeleteSelectedAsync(Func<int,bool> confirm)
    {
        if(busy)return 0;
        var selected=DraftGrid.SelectedItems.Cast<CampaignDraftEntry>().ToList();
        if(selected.Count==0||!confirm(selected.Count))return 0;
        busy=true;UpdateActions();RefreshButton.IsEnabled=false;
        int deleted=0;var failures=new List<string>();
        try
        {
            foreach(var entry in selected)
            {
                try{CampaignDraftLibrary.Delete(entry);deleted++;}
                catch(Exception ex){failures.Add(entry.Title+": "+ex.Message);}
            }
        }
        finally{busy=false;await RefreshAsync();}
        StatusText.Text=$"{deleted} draft(s) deleted. Attachments and source workbooks were kept."+(failures.Count>0?" Could not delete: "+string.Join("; ",failures):"");
        return deleted;
    }
    private async void Delete_Click(object sender,RoutedEventArgs e)=>await DeleteSelectedAsync(count=>MessageBox.Show(this,$"Delete {count} selected draft(s)?\n\nThis removes the saved drafts permanently. Attachments and source workbooks will be kept.","Delete email drafts",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)==MessageBoxResult.Yes);
    private void Close_Click(object sender,RoutedEventArgs e)=>Close();
}
