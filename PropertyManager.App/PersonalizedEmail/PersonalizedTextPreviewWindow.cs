using System.Windows;
using System.Windows.Controls;
namespace PropertyManager.App.PersonalizedEmail;
internal sealed class PersonalizedTextPreviewWindow:Window
{
    private readonly string template;
    private readonly List<(string Unit,Dictionary<string,string> Values)> rows;
    private readonly bool personalize;
    internal readonly Button PreviousButton=new(){Content="Previous",Padding=new Thickness(12,8,12,8),Margin=new Thickness(0,0,8,0)};
    internal readonly Button NextButton=new(){Content="Next",Padding=new Thickness(12,8,12,8),Margin=new Thickness(0,0,12,0)};
    internal readonly TextBlock Position=new(){VerticalAlignment=VerticalAlignment.Center};
    internal readonly TextBox Preview=new(){IsReadOnly=true,TextWrapping=TextWrapping.Wrap,VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Padding=new Thickness(20)};
    internal int Index {get;private set;}
    internal PersonalizedTextPreviewWindow(string template,IEnumerable<CampaignRecipient> recipients,bool personalize=true)
    {
        this.template=template;this.personalize=personalize;rows=recipients.Where(r=>r.Included).Select(r=>(r.Unit,r.Values())).ToList();
        if(rows.Count==0)throw new InvalidOperationException("Load and include at least one recipient first.");
        Width=800;Height=620;MinWidth=500;MinHeight=360;WindowStartupLocation=WindowStartupLocation.CenterOwner;
        var panel=new DockPanel{Margin=new Thickness(20)};var navigation=new WrapPanel{Margin=new Thickness(0,0,0,12)};
        navigation.Children.Add(PreviousButton);navigation.Children.Add(NextButton);navigation.Children.Add(Position);
        DockPanel.SetDock(navigation,Dock.Top);panel.Children.Add(navigation);panel.Children.Add(Preview);Content=panel;
        PreviousButton.Click+=(_,_)=>ShowRecord(Index-1);NextButton.Click+=(_,_)=>ShowRecord(Index+1);ShowRecord(0);
    }
    private void ShowRecord(int index)
    {
        if(index<0||index>=rows.Count)return;Index=index;var row=rows[index];
        Title="Personalized text preview - "+row.Unit;Position.Text=$"Record {index+1} of {rows.Count} - Unit {row.Unit}";
        PreviousButton.IsEnabled=index>0;NextButton.IsEnabled=index<rows.Count-1;
        try{Preview.Text="TEXT PREVIEW ONLY - Merge values use the selected record. Generate documents to check Word layout and formatting.\n\n"+(personalize?CampaignService.Render(template,row.Values):template);}
        catch(Exception ex){Preview.Text="This record cannot be previewed: "+ex.Message;}
        Preview.ScrollToHome();
    }
}
