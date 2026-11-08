using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using PropertyManager.App.EmailLists;
using Microsoft.Win32;
namespace PropertyManager.App.PersonalizedEmail;
public partial class RecipientDataEditorWindow:Window
{
    internal RecipientSnapshotSession Session {get;private set;}
    internal List<CampaignRecipient>? Result {get;private set;}
    internal RecipientDataEditorWindow(IEnumerable<CampaignRecipient> recipients)
    {
        InitializeComponent();Session=new(recipients);Bind();
    }
    private void Commit(){EditorGrid.CommitEdit(DataGridEditingUnit.Cell,true);EditorGrid.CommitEdit(DataGridEditingUnit.Row,true);}
    private void Bind(){EditorGrid.ItemsSource=Session.Table.DefaultView;CountText.Text=$"{Session.Table.Rows.Count} recipients | {Session.Table.Columns.Count-1} fields. New columns are available as personalization fields after Apply changes.";}
    private void Column_Generating(object sender,DataGridAutoGeneratingColumnEventArgs e)
    {
        if(e.PropertyName==RecipientSnapshotSession.Id){e.Cancel=true;return;}
        e.Column.Width=e.PropertyName=="Include"?85:e.PropertyName=="Unit"?90:200;
        if(e.Column is DataGridBoundColumn bound)bound.Binding=new Binding{Path=new PropertyPath("[(0)]",e.PropertyName),Mode=System.Windows.Data.BindingMode.TwoWay,UpdateSourceTrigger=UpdateSourceTrigger.LostFocus};
        e.Column.MinWidth=70;
    }
    private void Selection_Changed(object sender,SelectionChangedEventArgs e){if(EditRowButton is not null)EditRowButton.IsEnabled=EditorGrid.SelectedItem is DataRowView;}
    internal void AddColumn(string name){Commit();Session.AddColumn(name);EditorGrid.ItemsSource=null;Bind();}
    private void AddColumn_Click(object sender,RoutedEventArgs e)
    {
        var field=new TextBox{Margin=new Thickness(0,10,0,12),Padding=new Thickness(9),MinWidth=300};
        var help=new TextBlock{Text="Examples: Amount due, Parking space, Appointment date",TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,14)};
        var add=new Button{Content="Add column",HorizontalAlignment=HorizontalAlignment.Right};
        var stack=new StackPanel{Margin=new Thickness(24)};stack.Children.Add(new TextBlock{Text="Column name",FontSize=18,FontWeight=FontWeights.SemiBold});stack.Children.Add(field);stack.Children.Add(help);stack.Children.Add(add);
        var popup=FormWindow("Add personalization column",stack,460,300,this);
        add.Click+=(_,_)=>{try{AddColumn(field.Text);popup.DialogResult=true;}catch(Exception ex){help.Text=ex.Message;}};
        popup.ShowDialog();
    }
    private void Columns_Click(object sender,RoutedEventArgs e)
    {
        var menu=new ContextMenu();
        foreach(var column in EditorGrid.Columns)
        {
            var item=new MenuItem{Header=column.Header,IsCheckable=true,IsChecked=column.Visibility==Visibility.Visible,StaysOpenOnClick=true,IsEnabled=column.Header?.ToString() is not ("Unit" or "Names")};
            item.Click+=(_,_)=>column.Visibility=item.IsChecked?Visibility.Visible:Visibility.Collapsed;menu.Items.Add(item);
        }
        ColumnsButton.ContextMenu=menu;menu.PlacementTarget=ColumnsButton;menu.Placement=PlacementMode.Bottom;menu.IsOpen=true;
    }
    private static Window FormWindow(string title,object content,double width,double height,Window owner)=>new(){Title=title,Owner=owner,Width=width,Height=Math.Min(height,SystemParameters.WorkArea.Height-60),MinWidth=400,MinHeight=260,WindowStartupLocation=WindowStartupLocation.CenterOwner,Background=(Brush)Application.Current.FindResource("ThemePage"),Foreground=new SolidColorBrush(Color.FromRgb(23,43,58)),Content=content};
    internal bool EditRecord(string id,Window owner)
    {
        Commit();var row=Session.Table.Rows.Cast<DataRow>().Single(r=>(string)r[RecipientSnapshotSession.Id]==id);
        var grid=new Grid{Margin=new Thickness(24)};grid.RowDefinitions.Add(new(){Height=new GridLength(1,GridUnitType.Star)});grid.RowDefinitions.Add(new(){Height=GridLength.Auto});
        var fields=new StackPanel();var boxes=new Dictionary<string,Control>();
        foreach(DataColumn column in Session.Table.Columns)
        {
            string name=column.ColumnName;if(name==RecipientSnapshotSession.Id)continue;
            fields.Children.Add(new TextBlock{Text=name,Margin=new Thickness(0,8,0,5)});
            Control input=name=="Include"?new CheckBox{Content="Include this recipient",IsChecked=(bool)row[name]}:new TextBox{Text=Convert.ToString(row[name]),Padding=new Thickness(9),TextWrapping=TextWrapping.Wrap};
            System.Windows.Automation.AutomationProperties.SetName(input,name);boxes[name]=input;fields.Children.Add(input);
        }
        grid.Children.Add(new ScrollViewer{Content=fields,VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Padding=new Thickness(0,0,14,0)});
        var buttons=new StackPanel{Orientation=Orientation.Horizontal,HorizontalAlignment=HorizontalAlignment.Right,Margin=new Thickness(0,16,0,0)};
        var cancel=new Button{Content="Cancel",IsCancel=true,Margin=new Thickness(0,0,8,0)};var apply=new Button{Content="Apply recipient"};buttons.Children.Add(cancel);buttons.Children.Add(apply);Grid.SetRow(buttons,1);grid.Children.Add(buttons);
        var popup=FormWindow("Recipient - "+row["Unit"],grid,620,720,owner);
        apply.Click+=(_,_)=>{foreach(var (name,input) in boxes)row[name]=input is CheckBox check?check.IsChecked==true:((TextBox)input).Text;popup.DialogResult=true;};
        return popup.ShowDialog()==true;
    }
    private void EditRow_Click(object sender,RoutedEventArgs e){if(EditorGrid.SelectedItem is DataRowView row)EditRecord((string)row[RecipientSnapshotSession.Id],this);}
    private void Row_DoubleClick(object sender,MouseButtonEventArgs e){if(e.OriginalSource is DependencyObject source&&ItemsControl.ContainerFromElement(EditorGrid,source) is DataGridRow)EditRow_Click(sender,e);}
    internal void ExportSnapshot(string path)
    {
        // Never overwrite an original workbook, attachment, or existing export.
        string full=Path.GetFullPath(path);if(File.Exists(full))throw new IOException("Choose a new file name so existing workbooks remain unchanged.");string stage=full+"."+Guid.NewGuid().ToString("N")+".tmp";
        try{ExcelRecipientExporter.Export(stage,Session.Export());File.Move(stage,full,false);}
        finally{if(File.Exists(stage))File.Delete(stage);}
    }
    private void Export_Click(object sender,RoutedEventArgs e)
    {
        Commit();var picker=new SaveFileDialog{Title="Export recipient snapshot",Filter="Excel workbook (*.xlsx)|*.xlsx",DefaultExt=".xlsx",FileName="Recipient snapshot.xlsx",OverwritePrompt=false};if(picker.ShowDialog(this)!=true)return;
        try{ExportSnapshot(picker.FileName);StatusText.Text="Snapshot exported. Keep Recipient ID unchanged. Use Include = No to exclude a row; add columns as needed, then import the edited workbook.";}catch(Exception ex){StatusText.Text=ex.Message;}
    }
    private async void Import_Click(object sender,RoutedEventArgs e)
    {
        Commit();var picker=new OpenFileDialog{Title="Import edited recipient snapshot",Filter="Excel workbook (*.xlsx)|*.xlsx"};if(picker.ShowDialog(this)!=true)return;
        try
        {
            var loaded=await RecipientFileLoader.LoadAsync(picker.FileName,this,"Recipients");
            var imported=Session.Import(loaded.Data);
            if(MessageBox.Show(this,$"Apply {imported.ChangedCells} changed values, including {imported.ChangedEmails} email address changes?\n\nRows were matched by Recipient ID. Missing, unknown or duplicate IDs are rejected. Changes remain in this editor until Apply changes.","Import edited snapshot",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return;
            Session=imported.Session;Bind();StatusText.Text="Edited snapshot imported. Review email addresses and custom fields, then Apply changes.";
        }
        catch(OperationCanceledException){}
        catch(Exception ex){StatusText.Text=ex.Message;}
    }
    private void Apply_Click(object sender,RoutedEventArgs e){Commit();Result=Session.Recipients();DialogResult=true;}
}
