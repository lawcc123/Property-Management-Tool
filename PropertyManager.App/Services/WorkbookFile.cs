using System.Diagnostics;
using System.IO;
namespace PropertyManager.App.Services;
internal static class WorkbookFile
{
    internal static ProcessStartInfo Launch(string path,bool allowMacros)
    {
        if(string.IsNullOrWhiteSpace(path))throw new InvalidOperationException("Choose an Excel workbook using Browse files first.");
        string fullPath=Path.GetFullPath(path.Trim());
        string extension=Path.GetExtension(fullPath);
        if(!extension.Equals(".xlsx",StringComparison.OrdinalIgnoreCase)&&!(allowMacros&&extension.Equals(".xlsm",StringComparison.OrdinalIgnoreCase)))throw new InvalidOperationException("Choose a supported Excel workbook.");
        if(!File.Exists(fullPath))throw new FileNotFoundException("The workbook could not be found. Check its path and network connection.",fullPath);
        return new ProcessStartInfo(fullPath){UseShellExecute=true};
    }
}
