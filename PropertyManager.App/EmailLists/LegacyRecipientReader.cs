using System.Globalization;
using System.IO;
using System.Text;
using ExcelDataReader;

namespace PropertyManager.App.EmailLists;
internal static class LegacyRecipientReader
{
    internal static List<(string Sheet,RecipientFileData Data)> Read(string path)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var stream=File.Open(path,FileMode.Open,FileAccess.Read,FileShare.Read);
        using var reader=ExcelReaderFactory.CreateReader(stream);
        var sheets=new List<(string,RecipientFileData)>();
        do
        {
            if(!reader.Read())continue;
            var headers=Enumerable.Range(0,reader.FieldCount).Select(i=>(Index:i,Text:Convert.ToString(reader.GetValue(i),CultureInfo.InvariantCulture)?.Trim()??"")).Where(h=>h.Text.Length>0).ToList();
            if(headers.Count==0||headers.Select(h=>h.Text).Distinct(StringComparer.OrdinalIgnoreCase).Count()!=headers.Count)throw new InvalidDataException("The worksheet has missing or duplicate column headings.");
            var rows=new List<Dictionary<string,string>>();
            while(reader.Read())
            {
                var row=headers.ToDictionary(h=>h.Text,h=>Convert.ToString(reader.GetValue(h.Index),CultureInfo.InvariantCulture)?.Trim()??"",StringComparer.OrdinalIgnoreCase);
                if(row.Values.Any(v=>v.Length>0))rows.Add(row);
            }
            sheets.Add((reader.Name,new(headers.Select(h=>h.Text).ToList(),rows)));
        }while(reader.NextResult());
        return sheets;
    }
}
