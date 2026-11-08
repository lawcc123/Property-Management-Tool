using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace PropertyManager.App.EmailLists;

public partial class FolderShortcutEditWindow : Window
{
    private readonly List<FolderShortcut> _otherFolders;
    internal FolderShortcut? SavedFolder { get; private set; }

    internal FolderShortcutEditWindow(FolderShortcut folder, IEnumerable<FolderShortcut> otherFolders, bool editing = false)
    {
        _otherFolders = otherFolders.ToList();
        InitializeComponent();
        Title = HeadingText.Text = "Add/Edit Condo Corporation";
        NameCombo.ItemsSource=BuildingLink.CorporationConfiguration.Names.Where(n=>!_otherFolders.Any(f=>f.Name==n)).ToList();
        NameCombo.SelectedItem=BuildingLink.CorporationConfiguration.Canonical(folder.Name);if(NameCombo.SelectedItem is null)NameCombo.SelectedIndex=0;
        NameCombo.IsEnabled=!editing;
        PathText.Text = folder.Path;
        UnitPathText.Text=folder.UnitPath??"";
        ColorCombo.ItemsSource = FolderColors.Options;
        ColorCombo.SelectedValue = string.IsNullOrWhiteSpace(folder.Name)?FolderColors.Resolve(null,NameCombo.SelectedItem as string??""):folder.ButtonColor;
        Loaded += (_, _) => NameCombo.Focus();
        NameCombo.SelectionChanged+=(_,_)=>{if(!editing)ColorCombo.SelectedValue=FolderColors.Resolve(null,NameCombo.SelectedItem as string??"");};
    }

    private void IgnoreWheel(object s,System.Windows.Input.MouseWheelEventArgs e){if(s is System.Windows.Controls.ComboBox{IsDropDownOpen:false})e.Handled=true;}
    private void Browse_Click(object sender, RoutedEventArgs e)
    {
        OpenFolderDialog picker = new() { Title = "Choose BuildingLink occupant data backup folder", Multiselect = false };
        if (picker.ShowDialog(this) == true) PathText.Text = picker.FolderName;
    }

    private void BrowseUnit_Click(object sender,RoutedEventArgs e)
    {
        var picker=new OpenFolderDialog{Title="Choose BuildingLink Unit Data backup folder",Multiselect=false};
        if(picker.ShowDialog(this)==true)UnitPathText.Text=picker.FolderName;
    }
    private void Save_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            SavedFolder = FolderShortcutStore.ValidateEntry(NameCombo.SelectedItem as string??"", PathText.Text, _otherFolders, ColorCombo.SelectedValue as string,UnitPathText.Text);
            DialogResult = true;
        }
        catch (Exception ex) when (ex is InvalidDataException or ArgumentException or NotSupportedException)
        {
            ErrorText.Text = ex.Message;
            ErrorText.Visibility = Visibility.Visible;
        }
    }
}
