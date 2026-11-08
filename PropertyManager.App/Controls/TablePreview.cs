using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace PropertyManager.App.Controls;

// Move the existing table, rather than rebuilding columns or copying editable rows.
public static class TablePreview
{
    public static readonly DependencyProperty TargetProperty=DependencyProperty.RegisterAttached("Target",typeof(DataGrid),typeof(TablePreview),new PropertyMetadata(null,TargetChanged));
    public static readonly DependencyProperty TitleProperty=DependencyProperty.RegisterAttached("Title",typeof(string),typeof(TablePreview),new PropertyMetadata("Enlarged table preview"));
    public static void SetTarget(DependencyObject element,DataGrid value)=>element.SetValue(TargetProperty,value);
    public static DataGrid? GetTarget(DependencyObject element)=>(DataGrid?)element.GetValue(TargetProperty);
    public static void SetTitle(DependencyObject element,string value)=>element.SetValue(TitleProperty,value);
    public static string GetTitle(DependencyObject element)=>(string)element.GetValue(TitleProperty);
    private static void TargetChanged(DependencyObject element,DependencyPropertyChangedEventArgs args)
    {
        if(element is Button button){button.Click-=Open;button.Click+=Open;}
    }
    private static void Open(object sender,RoutedEventArgs args)
    {
        if(sender is not Button button||GetTarget(button) is not {} table)return;
        var owner=Window.GetWindow(button);
        Window? expanded=null;
        try{expanded=Create(table,owner,GetTitle(button));expanded.ShowDialog();}
        catch(Exception ex){expanded?.Close();MessageBox.Show(owner,ex.Message,"Enlarge preview",MessageBoxButton.OK,MessageBoxImage.Warning);}
    }
    internal static Window Create(DataGrid table,Window? owner,string title)
    {
        if(table.Parent is not Panel parent)throw new InvalidOperationException("This table cannot be enlarged from its current location.");
        if(!table.CommitEdit(DataGridEditingUnit.Cell,true)||!table.CommitEdit(DataGridEditingUnit.Row,true))throw new InvalidOperationException("Finish or correct the current table edit before enlarging it.");
        int index=parent.Children.IndexOf(table);
        DependencyProperty[] properties=[FrameworkElement.HeightProperty,FrameworkElement.WidthProperty,FrameworkElement.MaxHeightProperty,FrameworkElement.MaxWidthProperty,FrameworkElement.MinHeightProperty,FrameworkElement.MinWidthProperty,FrameworkElement.VerticalAlignmentProperty,FrameworkElement.HorizontalAlignmentProperty,FrameworkElement.MarginProperty,UIElement.VisibilityProperty];
        var original=properties.Select(p=>(Property:p,Value:table.ReadLocalValue(p),Binding:BindingOperations.GetBindingBase(table,p))).ToList();
        var window=new Window{Owner=owner,Title=title,Width=Math.Min(1240,SystemParameters.WorkArea.Width),Height=Math.Min(820,SystemParameters.WorkArea.Height),MinWidth=640,MinHeight=400,WindowStartupLocation=WindowStartupLocation.CenterOwner,Background=new SolidColorBrush(Color.FromRgb(244,247,249)),Foreground=new SolidColorBrush(Color.FromRgb(23,43,58))};
        var root=new DockPanel{Margin=new Thickness(24)};window.Content=root;
        var heading=new DockPanel{Margin=new Thickness(0,0,0,16)};DockPanel.SetDock(heading,Dock.Top);root.Children.Add(heading);
        var close=new Button{Content="Back to table",Padding=new Thickness(12,8,12,8),Margin=new Thickness(12,0,0,0)};DockPanel.SetDock(close,Dock.Right);heading.Children.Add(close);close.Click+=(_,_)=>window.Close();
        heading.Children.Add(new TextBlock{Text=title,FontSize=22,FontWeight=FontWeights.SemiBold,TextWrapping=TextWrapping.Wrap,VerticalAlignment=VerticalAlignment.Center});
        var hint=new TextBlock{Text="Uses the current rows, filters and selections. Edits remain available when you return. Close this window to use the feature's actions.",TextWrapping=TextWrapping.Wrap,Foreground=Brushes.SlateGray,Margin=new Thickness(0,0,0,12)};DockPanel.SetDock(hint,Dock.Top);root.Children.Add(hint);
        parent.Children.Remove(table);
        table.Height=table.Width=double.NaN;table.MaxHeight=table.MaxWidth=double.PositiveInfinity;table.MinHeight=table.MinWidth=0;table.VerticalAlignment=VerticalAlignment.Stretch;table.HorizontalAlignment=HorizontalAlignment.Stretch;table.Margin=new Thickness(0);table.Visibility=Visibility.Visible;
        root.Children.Add(table);bool restored=false;
        void Restore()
        {
            if(restored)return;restored=true;root.Children.Remove(table);
            foreach(var saved in original)
            {
                if(saved.Binding is not null)BindingOperations.SetBinding(table,saved.Property,saved.Binding);
                else if(saved.Value==DependencyProperty.UnsetValue)table.ClearValue(saved.Property);
                else table.SetValue(saved.Property,saved.Value);
            }
            parent.Children.Insert(Math.Min(index,parent.Children.Count),table);
        }
        window.Closing+=(_,args)=>{if(!table.CommitEdit(DataGridEditingUnit.Cell,true)||!table.CommitEdit(DataGridEditingUnit.Row,true)){args.Cancel=true;MessageBox.Show(window,"Correct the current edit before returning to the table.","Table edit");}};
        window.Loaded+=(_,_)=>
        {
            // Recalculate star widths after moving from a constrained page layout.
            foreach(var column in table.Columns.Where(c=>c.Width.IsStar))
            {var width=column.Width;column.Width=DataGridLength.Auto;column.Width=new DataGridLength(width.Value,DataGridLengthUnitType.Star);}
            table.UpdateLayout();
        };
        window.Closed+=(_,_)=>Restore();
        return window;
    }
}
