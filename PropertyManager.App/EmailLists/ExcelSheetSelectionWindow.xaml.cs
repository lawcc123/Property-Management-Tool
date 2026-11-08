using System.Windows;

namespace PropertyManager.App.EmailLists;

public partial class ExcelSheetSelectionWindow : Window
{
    public ExcelWorksheetInfo? SelectedWorksheet { get; private set; }

    public ExcelSheetSelectionWindow(
        List<ExcelWorksheetInfo> worksheets,
        string? preferredWorksheetName)
    {
        InitializeComponent();
        WorksheetsListBox.ItemsSource = worksheets;

        ExcelWorksheetInfo? preferredWorksheet = worksheets.FirstOrDefault(worksheet =>
            string.Equals(worksheet.Name, preferredWorksheetName, StringComparison.OrdinalIgnoreCase));

        WorksheetsListBox.SelectedItem = preferredWorksheet ?? worksheets.FirstOrDefault();
    }

    private void UseSheet_Click(object sender, RoutedEventArgs e)
    {
        if (WorksheetsListBox.SelectedItem is not ExcelWorksheetInfo worksheet)
        {
            MessageBox.Show(
                this,
                "Please select one worksheet.",
                "No Worksheet Selected",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        SelectedWorksheet = worksheet;
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
