using System.Globalization;
using System.Windows;
using System.Windows.Data;
using PropertyManager.App.EmailLists;

namespace PropertyManager.App.OutlookSync;

public sealed class OutlookPageHeightConverter : IMultiValueConverter
{
    private readonly EmailPageHeightConverter sizing=new();
    public object Convert(object[] values,Type targetType,object parameter,CultureInfo culture)
    {
        // Collapsed WPF controls can retain their previous ActualHeight. That
        // old setup-card height must not reserve blank space in history mode.
        var effective=values.ToArray();
        if(effective.Length>3&&effective[3] is Visibility visibility&&visibility!=Visibility.Visible)effective[2]=0.0;
        return sizing.Convert(effective,targetType,parameter,culture);
    }
    public object[] ConvertBack(object value,Type[] targetTypes,object parameter,CultureInfo culture)=>throw new NotSupportedException();
}
