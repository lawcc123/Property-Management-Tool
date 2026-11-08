using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace PropertyManager.App.Envelopes;

internal sealed class LabelPositionsWindow : Window
{
    private readonly LabelLayout sheet;
    private readonly RadioButton startMode=new(){Content="Start at position",GroupName="PositionMode",VerticalAlignment=VerticalAlignment.Center};
    private readonly RadioButton specificMode=new(){Content="Print only in the positions I check below",GroupName="PositionMode",Margin=new Thickness(0,14,0,8)};
    private readonly ComboBox start=new(){MinWidth=70,Padding=new Thickness(7),Margin=new Thickness(12,0,0,0)};
    private readonly List<CheckBox> positions=[];
    private readonly TextBlock error=new(){Foreground=Brushes.Firebrick,TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,10,0,0)};
    internal int FirstSlot {get;private set;}=1;
    internal List<int>? SpecificSlots {get;private set;}
    internal LabelPositionsWindow(LabelLayout sheet,int first,List<int>? selected)
    {
        this.sheet=sheet;Title="Choose labels to print";Width=600;Height=700;MinWidth=480;MinHeight=500;
        WindowStartupLocation=WindowStartupLocation.CenterOwner;Background=new SolidColorBrush(Color.FromRgb(244,247,249));
        Foreground=new SolidColorBrush(Color.FromRgb(23,43,58));
        var root=new Grid{Margin=new Thickness(24)};root.RowDefinitions.Add(new(){Height=GridLength.Auto});root.RowDefinitions.Add(new(){Height=new GridLength(1,GridUnitType.Star)});root.RowDefinitions.Add(new(){Height=GridLength.Auto});Content=root;
        var header=new StackPanel();root.Children.Add(header);
        header.Children.Add(new TextBlock{Text="Choose labels to print",FontSize=24,FontWeight=FontWeights.SemiBold});
        header.Children.Add(new TextBlock{Text=sheet.Label,TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,8,0,14)});
        var content=new StackPanel();var scroll=new ScrollViewer{Content=content,VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Padding=new Thickness(0,0,12,0)};Grid.SetRow(scroll,1);root.Children.Add(scroll);
        var startRow=new StackPanel{Orientation=Orientation.Horizontal};startRow.Children.Add(startMode);startRow.Children.Add(start);content.Children.Add(startRow);
        start.ItemsSource=Enumerable.Range(1,sheet.Capacity).ToList();start.SelectedItem=first;startMode.IsChecked=selected is null;specificMode.IsChecked=selected is not null;
        content.Children.Add(specificMode);
        var buttons=new WrapPanel{Margin=new Thickness(0,0,0,10)};
        var all=new Button{Content="Select all",Padding=new Thickness(12,6,12,6),Margin=new Thickness(0,0,8,0)};
        var clear=new Button{Content="Clear positions",Padding=new Thickness(12,6,12,6)};buttons.Children.Add(all);buttons.Children.Add(clear);content.Children.Add(buttons);
        var grid=new UniformGrid{Columns=sheet.Columns};
        for(int number=1;number<=sheet.Capacity;number++)
        {
            var position=new CheckBox{Content="Position "+number,Tag=number,IsChecked=selected?.Contains(number)==true,Margin=new Thickness(5,7,10,7)};
            position.Click+=(_,_)=>specificMode.IsChecked=true;positions.Add(position);grid.Children.Add(position);
        }
        content.Children.Add(new Border{Background=Brushes.White,BorderBrush=new SolidColorBrush(Color.FromRgb(220,229,233)),BorderThickness=new Thickness(1),CornerRadius=new CornerRadius(8),Padding=new Thickness(10),Child=grid});
        content.Children.Add(new TextBlock{Text="Numbered left to right, from the top row. These choices apply to the first sheet; additional sheets use all positions. One address goes in each position. Increase Copies per recipient to repeat an address. Position numbers and outlines are never printed.",TextWrapping=TextWrapping.Wrap,FontSize=12,Margin=new Thickness(0,12,0,0)});
        content.Children.Add(error);
        all.Click+=(_,_)=>{specificMode.IsChecked=true;foreach(var p in positions)p.IsChecked=true;};
        clear.Click+=(_,_)=>{specificMode.IsChecked=true;foreach(var p in positions)p.IsChecked=false;};
        startMode.Checked+=(_,_)=>start.IsEnabled=true;specificMode.Checked+=(_,_)=>start.IsEnabled=false;start.IsEnabled=selected is null;
        var footer=new StackPanel{Orientation=Orientation.Horizontal,HorizontalAlignment=HorizontalAlignment.Right,Margin=new Thickness(0,18,0,0)};Grid.SetRow(footer,2);root.Children.Add(footer);
        var cancel=new Button{Content="Cancel",IsCancel=true,Padding=new Thickness(14,8,14,8),Margin=new Thickness(0,0,8,0)};
        var apply=new Button{Content="Use selected positions",IsDefault=true,Padding=new Thickness(14,8,14,8),Background=new SolidColorBrush(Color.FromRgb(20,125,112)),Foreground=Brushes.White};
        apply.SetResourceReference(System.Windows.Controls.Control.BackgroundProperty,"ThemeAccent");apply.SetResourceReference(System.Windows.Controls.Control.ForegroundProperty,"ThemeAccentForeground");
        apply.Click+=(_,_)=>{if(TryApply())DialogResult=true;};footer.Children.Add(cancel);footer.Children.Add(apply);
    }
    internal bool TryApply()
    {
        FirstSlot=start.SelectedItem is int first?first:1;
        SpecificSlots=specificMode.IsChecked==true?positions.Where(p=>p.IsChecked==true).Select(p=>(int)p.Tag).ToList():null;
        if(SpecificSlots is {Count:0}){error.Text="Check at least one label position, or choose Start at position.";return false;}
        return true;
    }
}
