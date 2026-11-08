using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace PropertyManager.App.Envelopes;

internal sealed class EnvelopePreviewWindow : Window
{
    internal EnvelopePreviewWindow(EnvelopeDrawing drawing,string caption,bool fitPage=false)
    {
        Title="Mailing preview";Width=1080;Height=720;MinWidth=700;MinHeight=500;WindowStartupLocation=WindowStartupLocation.CenterOwner;
        Background=new SolidColorBrush(Color.FromRgb(244,247,249));Foreground=new SolidColorBrush(Color.FromRgb(23,43,58));
        var grid=new Grid{Margin=new Thickness(24)};grid.RowDefinitions.Add(new(){Height=GridLength.Auto});grid.RowDefinitions.Add(new(){Height=GridLength.Auto});grid.RowDefinitions.Add(new(){Height=new GridLength(1,GridUnitType.Star)});grid.RowDefinitions.Add(new(){Height=GridLength.Auto});
        var heading=new TextBlock{Text=caption,FontSize=22,FontWeight=FontWeights.SemiBold,TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,14)};grid.Children.Add(heading);
        var tools=new StackPanel{Orientation=Orientation.Horizontal,Margin=new Thickness(0,0,0,14)};Grid.SetRow(tools,1);grid.Children.Add(tools);
        var zoom=new Slider{Minimum=10,Maximum=200,Value=100,Width=230,VerticalAlignment=VerticalAlignment.Center,Margin=new Thickness(12,0,12,0)};
        var amount=new TextBlock{Text="100%",Width=58,VerticalAlignment=VerticalAlignment.Center};
        tools.Children.Add(new TextBlock{Text="Zoom",VerticalAlignment=VerticalAlignment.Center});tools.Children.Add(zoom);tools.Children.Add(amount);
        var fit=new Button{Content="Fit page",Padding=new Thickness(12,7,12,7)};tools.Children.Add(fit);
        var image=new Image{Source=new DrawingImage(drawing.Drawing),Width=drawing.Page.Width,Height=drawing.Page.Height,Stretch=Stretch.Fill,HorizontalAlignment=HorizontalAlignment.Center,VerticalAlignment=VerticalAlignment.Center,Margin=new Thickness(18)};
        var scroll=new ScrollViewer{Content=image,HorizontalScrollBarVisibility=ScrollBarVisibility.Auto,VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Background=new SolidColorBrush(Color.FromRgb(228,235,240))};Grid.SetRow(scroll,2);grid.Children.Add(scroll);
        zoom.ValueChanged+=(_,_)=>{image.Width=drawing.Page.Width*zoom.Value/100;image.Height=drawing.Page.Height*zoom.Value/100;amount.Text=$"{zoom.Value:0}%";};
        fit.Click+=(_,_)=>zoom.Value=Math.Clamp(Math.Min((scroll.ViewportWidth-40)/drawing.Page.Width,(scroll.ViewportHeight-40)/drawing.Page.Height)*100,10,200);
        if(fitPage)Loaded+=(_,_)=>fit.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        var footer=new TextBlock{Text="Same layout as direct printing. Dashed guides are not printed. Screen zoom does not change the paper's physical dimensions.",TextWrapping=TextWrapping.Wrap,FontSize=12,Margin=new Thickness(0,12,0,0)};Grid.SetRow(footer,3);grid.Children.Add(footer);Content=grid;
    }
}
