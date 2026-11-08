using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;

namespace PropertyManager.App.EmailLists;

public partial class ManageFoldersWindow : Window
{
    private readonly ObservableCollection<FolderShortcut> _folders;
    private readonly Dictionary<string, string> _pathChanges = new(StringComparer.OrdinalIgnoreCase);
    internal List<FolderShortcut> SavedFolders => _folders.ToList();
    internal IReadOnlyDictionary<string, string> PathChanges => _pathChanges;

    internal ManageFoldersWindow(IEnumerable<FolderShortcut> folders)
    {
        _folders = new(folders);
        InitializeComponent();
        FoldersGrid.ItemsSource = _folders;
        if (_folders.Count > 0) FoldersGrid.SelectedIndex = 0;
        UpdateActions();
    }

    private void Selection_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (IsInitialized) UpdateActions();
    }
    private void UpdateActions()
    {
        RenameButton.IsEnabled=ChangePathButton.IsEnabled=RemoveButton.IsEnabled=FoldersGrid.SelectedItem is FolderShortcut;
        AddButton.IsEnabled=BuildingLink.CorporationConfiguration.Names.Any(n=>!_folders.Any(f=>f.Name==n));
    }

    private void Edit(FolderShortcut folder, string? newPath = null)
    {
        FolderShortcutEditWindow editor = new(folder with { Path = newPath ?? folder.Path }, _folders.Where(f => f != folder), true) { Owner = this };
        if (editor.ShowDialog() != true || editor.SavedFolder is not { } updated) return;
        int index = _folders.IndexOf(folder);
        foreach (string key in _pathChanges.Where(pair => pair.Value.Equals(folder.Path, StringComparison.OrdinalIgnoreCase)).Select(pair => pair.Key).ToList())
            _pathChanges[key] = updated.Path;
        _pathChanges[folder.Path] = updated.Path;
        _folders[index] = updated;
        FoldersGrid.SelectedItem = updated;UpdateActions();
    }
    private void Rename_Click(object sender, RoutedEventArgs e)
    {
        if (FoldersGrid.SelectedItem is FolderShortcut folder) Edit(folder);
    }
    private void ChangePath_Click(object sender, RoutedEventArgs e)
    {
        if (FoldersGrid.SelectedItem is FolderShortcut folder) Edit(folder);
    }
    private void Add_Click(object sender, RoutedEventArgs e)
    {
        FolderShortcutEditWindow editor = new(new("", ""), _folders) { Owner = this };
        if (editor.ShowDialog() == true && editor.SavedFolder is { } added) { _folders.Add(added); FoldersGrid.SelectedItem = added;UpdateActions(); }
    }
    private void Remove_Click(object sender, RoutedEventArgs e)
    {
        if (FoldersGrid.SelectedItem is FolderShortcut folder)
        {if(MessageBox.Show(this,"Remove "+folder.Name+" from BuildingLink downloads, email lists, owner mailing lists and Outlook sync? Existing files and Outlook contacts will remain unchanged.","Remove shared corporation",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)!=MessageBoxResult.Yes)return;_folders.Remove(folder);}
        if (_folders.Count > 0) FoldersGrid.SelectedIndex = 0;
        UpdateActions();
    }
    private void Save_Click(object sender,RoutedEventArgs e)
    {
        try
        {
            foreach(var f in _folders){if(!BuildingLink.CorporationConfiguration.Names.Contains(f.Name))throw new System.IO.InvalidDataException("Choose DEMO 1001 or DEMO 2002.");FolderShortcutStore.ValidateEntry(f.Name,f.Path,_folders.Where(o=>o!=f),f.Color,f.UnitPath);}
            DialogResult=true;
        }catch(Exception ex){MessageBox.Show(this,ex.Message,"Manage Condo Corporations",MessageBoxButton.OK,MessageBoxImage.Warning);}
    }
}
