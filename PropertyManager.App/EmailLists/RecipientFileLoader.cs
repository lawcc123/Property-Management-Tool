using System.IO;
using System.Windows;

namespace PropertyManager.App.EmailLists;

internal static class RecipientFileLoader
{
    internal static async Task<(RecipientFileData Data, string Sheet)> LoadAsync(string path, Window owner, string preferred)
    {
        if (Path.GetExtension(path).Equals(".csv",StringComparison.OrdinalIgnoreCase))
            return (await Task.Run(() => CsvRecipientReader.Read(path)), "CSV");
        if(Path.GetExtension(path).Equals(".xls",StringComparison.OrdinalIgnoreCase))
        {
            List<(string Sheet,RecipientFileData Data)> legacy;
            try{legacy=await Task.Run(()=>LegacyRecipientReader.Read(path));}
            catch(Exception ex){throw new InvalidDataException("The legacy Excel file could not be read. Choose a fresh BuildingLink Unit Data export.",ex);}
            if(legacy.Count==0)throw new InvalidDataException("The workbook contains no worksheets.");
            int index=0;
            if(legacy.Count>1)
            {
                var choices=legacy.Select((s,n)=>new ExcelWorksheetInfo(s.Sheet,n.ToString())).ToList();
                var picker=new ExcelSheetSelectionWindow(choices,preferred){Owner=owner};
                if(picker.ShowDialog()!=true||picker.SelectedWorksheet is null)throw new OperationCanceledException();
                index=int.Parse(picker.SelectedWorksheet.WorksheetPath);
            }
            return (legacy[index].Data,legacy[index].Sheet);
        }
        var sheets = await Task.Run(() => ExcelRecipientReader.ListWorksheets(path));
        if (sheets.Count == 0) throw new InvalidDataException("The workbook contains no worksheets.");
        var chosen = sheets[0];
        if (sheets.Count > 1)
        {
            ExcelSheetSelectionWindow picker = new(sheets,preferred) { Owner = owner };
            if (picker.ShowDialog() != true || picker.SelectedWorksheet is null) throw new OperationCanceledException();
            chosen = picker.SelectedWorksheet;
        }
        return (await Task.Run(() => ExcelRecipientReader.ReadSheet(path,chosen)), chosen.Name);
    }
}
