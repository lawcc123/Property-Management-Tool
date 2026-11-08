using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace PropertyManager.App.EmailLists;

public partial class GeneralEmailListView
{
    private readonly List<string> _selectedUnits = [];

    internal static List<string> MatchingUnits(RecipientFileData? source, string query, IEnumerable<string> selected)
    {
        query = query.Trim();
        if (source is null || query.Length == 0) return [];
        var excluded = new HashSet<string>(selected, StringComparer.OrdinalIgnoreCase);
        return source.Rows.Select(r => GeneratedEmailListService.Value(r, "Unit ID"))
            .Where(unit => unit.Length > 0 && unit.StartsWith(query, StringComparison.OrdinalIgnoreCase) && !excluded.Contains(unit))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(unit => unit.Equals(query, StringComparison.OrdinalIgnoreCase) ? 0 : 1)
            .ThenBy(unit => long.TryParse(unit, out long number) ? number : long.MaxValue)
            .ThenBy(unit => unit, StringComparer.OrdinalIgnoreCase).ToList();
    }
    private void ResetUnitSelection()
    {
        _selectedUnits.Clear();
        UnitSearch.Clear();
        RenderSelectedUnits();
        RefreshUnitSuggestions();
    }
    private void UnitSearch_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (!IsInitialized) return;
        RefreshUnitSuggestions();
        if (!_updating && Mode == EmailListMode.Units) MarkFiltersChanged();
    }
    private void RefreshUnitSuggestions()
    {
        var matches = MatchingUnits(_source, UnitSearch.Text, _selectedUnits);
        UnitSuggestions.ItemsSource = matches;
        UnitSuggestions.SelectedIndex = matches.Count > 0 ? 0 : -1;
        UnitSuggestions.Visibility = matches.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        UnitMatchText.Text = string.IsNullOrWhiteSpace(UnitSearch.Text) ? "" : matches.Count == 0
            ? "No matching unselected units in this backup." : $"{matches.Count} matching units";
    }
    private void RenderSelectedUnits()
    {
        SelectedUnitTags.Children.Clear();
        foreach (string unit in _selectedUnits)
        {
            var tag = new Button { Content = unit + "  \u00D7", Padding = new Thickness(9,5,9,5),
                Margin = new Thickness(0,0,6,6), Background = new SolidColorBrush(Color.FromRgb(232,245,242)),
                Foreground = new SolidColorBrush(Color.FromRgb(20,125,112)), ToolTip = "Remove unit " + unit };
            System.Windows.Automation.AutomationProperties.SetName(tag, "Remove unit " + unit);
            tag.Click += (_,_) => { _selectedUnits.Remove(unit); RenderSelectedUnits(); RefreshUnitSuggestions(); MarkFiltersChanged(); };
            SelectedUnitTags.Children.Add(tag);
        }
        NoUnitsText.Visibility = _selectedUnits.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }
    private void AddSuggestedUnit()
    {
        string? unit = UnitSuggestions.SelectedItem as string ?? UnitSuggestions.Items.Cast<string>().FirstOrDefault();
        if (unit is null) return;
        if (!_selectedUnits.Contains(unit, StringComparer.OrdinalIgnoreCase)) _selectedUnits.Add(unit);
        UnitSearch.Clear();
        RenderSelectedUnits();
        RefreshUnitSuggestions();
        MarkFiltersChanged();
        UnitSearch.Focus();
    }
    private void Units_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        // Leave tag buttons' Enter/Space behavior intact for keyboard removal.
        if (!UnitSearch.IsKeyboardFocusWithin && !UnitSuggestions.IsKeyboardFocusWithin) return;
        if (e.Key == Key.Escape) { UnitSearch.Clear(); e.Handled = true; return; }
        if (UnitSuggestions.Items.Count == 0) return;
        if (e.Key is Key.Right or Key.Down or Key.Up || e.Key == Key.Left && UnitSuggestions.IsKeyboardFocusWithin)
        {
            int direction = e.Key is Key.Up or Key.Left ? -1 : 1;
            int count = UnitSuggestions.Items.Count;
            UnitSuggestions.SelectedIndex = (UnitSuggestions.SelectedIndex + direction + count) % count;
            UnitSuggestions.ScrollIntoView(UnitSuggestions.SelectedItem);
            e.Handled = true;
        }
        else if (e.Key == Key.Enter) { AddSuggestedUnit(); e.Handled = true; }
    }
    private void UnitSuggestions_MouseUp(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton == MouseButton.Left && ItemsControl.ContainerFromElement(UnitSuggestions, e.OriginalSource as DependencyObject) is ListBoxItem)
        { AddSuggestedUnit(); e.Handled = true; }
    }
}
