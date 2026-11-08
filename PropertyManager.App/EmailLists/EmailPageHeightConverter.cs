using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace PropertyManager.App.EmailLists;

public sealed class EmailPageHeightConverter : IMultiValueConverter
{
    public object Convert(object[] values,Type targetType,object parameter,CultureInfo culture)
    {
        double Size(int index)=>values.Length>index&&values[index] is double size&&double.IsFinite(size)?size:0;
        // Page margin: 27 + 24. Card gaps: 20 + 18. Reserve real working space
        // for both lower cards; overflow belongs to the page, not their contents.
        return Math.Max(Math.Max(0,Size(0)-51),Size(1)+Size(2)+38+420);
    }
    public object[] ConvertBack(object value,Type[] targetTypes,object parameter,CultureInfo culture)=>throw new NotSupportedException();
}
