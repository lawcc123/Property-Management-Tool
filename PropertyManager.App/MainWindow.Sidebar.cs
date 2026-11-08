using System.Windows;
using System.Windows.Controls;

namespace PropertyManager.App;

public partial class MainWindow
{
    private void SidebarScroll_Changed(object sender,ScrollChangedEventArgs e)
    {
        if(SidebarTopFade is null||SidebarBottomFade is null)return;
        SidebarTopFade.Visibility=SidebarScroll.VerticalOffset>0.5?Visibility.Visible:Visibility.Collapsed;
        SidebarBottomFade.Visibility=SidebarScroll.VerticalOffset<SidebarScroll.ScrollableHeight-0.5?Visibility.Visible:Visibility.Collapsed;
    }
}
