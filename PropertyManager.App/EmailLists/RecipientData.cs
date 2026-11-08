using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace PropertyManager.App.EmailLists;

public sealed record RecipientFileData(
    List<string> Columns,
    List<Dictionary<string, string>> Rows);

internal static class CsvRecipientReader
{
    public static RecipientFileData Read(string path)
    {
        using var reader = new Microsoft.VisualBasic.FileIO.TextFieldParser(path, Encoding.UTF8, true);
        reader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
        reader.SetDelimiters(",");
        reader.HasFieldsEnclosedInQuotes = true;
        reader.TrimWhiteSpace = false;
        if (reader.EndOfData) return new([], []);
        var headers = (reader.ReadFields() ?? []).Select(h => h.Trim()).ToList();
        if (headers.Any(h => h.Length == 0) || headers.Distinct(StringComparer.OrdinalIgnoreCase).Count() != headers.Count)
            throw new InvalidDataException("The CSV header contains blank or duplicate columns. Use a fresh BuildingLink export.");
        List<Dictionary<string,string>> rows = [];
        while (!reader.EndOfData)
        {
            var fields = reader.ReadFields() ?? [];
            if (fields.Length > headers.Count) throw new InvalidDataException("A CSV row contains more values than its header.");
            var row = headers.Select((h,n) => (h, value:n < fields.Length ? fields[n].Trim() : ""))
                .ToDictionary(p => p.h, p => p.value, StringComparer.OrdinalIgnoreCase);
            if (row.Values.Any(v => v.Length > 0)) rows.Add(row);
        }
        return new(headers,rows);
    }
}

internal static class ExcelRecipientReader
{
    private static readonly XNamespace SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationshipsNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    public static List<ExcelWorksheetInfo> ListWorksheets(string workbookPath)
    {
        using ZipArchive archive = ZipFile.OpenRead(workbookPath);

        ZipArchiveEntry workbookEntry = archive.GetEntry("xl/workbook.xml")
            ?? throw new InvalidDataException("The workbook is missing xl/workbook.xml.");
        ZipArchiveEntry relationshipsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
            ?? throw new InvalidDataException("The workbook is missing xl/_rels/workbook.xml.rels.");

        XDocument workbook = ReadXml(workbookEntry);
        XDocument relationships = ReadXml(relationshipsEntry);

        Dictionary<string, string> worksheetPathsByRelationshipId = relationships
            .Descendants(PackageRelationshipsNs + "Relationship")
            .Where(relationship =>
                string.Equals(
                    (string?)relationship.Attribute("Type"),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                    StringComparison.OrdinalIgnoreCase))
            .ToDictionary(
                relationship => (string)relationship.Attribute("Id")!,
                relationship => NormalizeWorkbookPartPath((string?)relationship.Attribute("Target") ?? string.Empty));

        return workbook
            .Descendants(SpreadsheetNs + "sheet")
            .Select(sheet =>
            {
                string name = (string?)sheet.Attribute("name") ?? "Worksheet";
                string relationshipId = (string?)sheet.Attribute(RelationshipsNs + "id") ?? string.Empty;
                return worksheetPathsByRelationshipId.TryGetValue(relationshipId, out string? worksheetPath)
                    ? new ExcelWorksheetInfo(name, worksheetPath)
                    : null;
            })
            .OfType<ExcelWorksheetInfo>()
            .ToList();
    }

    public static RecipientFileData ReadFirstSheet(string workbookPath)
    {
        List<ExcelWorksheetInfo> worksheets = ListWorksheets(workbookPath);
        if (worksheets.Count == 0)
        {
            return new RecipientFileData([], []);
        }

        return ReadSheet(workbookPath, worksheets[0]);
    }

    public static RecipientFileData ReadSheet(string workbookPath, ExcelWorksheetInfo worksheet)
    {
        using ZipArchive archive = ZipFile.OpenRead(workbookPath);

        List<string> sharedStrings = ReadSharedStrings(archive);
        string worksheetPath = worksheet.WorksheetPath;
        ZipArchiveEntry worksheetEntry = archive.GetEntry(worksheetPath)
            ?? throw new InvalidDataException($"Worksheet part was not found: {worksheetPath}");

        XDocument worksheetXml = ReadXml(worksheetEntry);
        XElement? firstRow = worksheetXml
            .Descendants(SpreadsheetNs + "sheetData")
            .Elements(SpreadsheetNs + "row")
            .FirstOrDefault();

        if (firstRow is null)
        {
            return new RecipientFileData([], []);
        }

        var headers = firstRow.Elements(SpreadsheetNs + "c")
            .Select(cell => new KeyValuePair<int, string>(
                GetColumnIndex((string?)cell.Attribute("r")), ReadCellText(cell, sharedStrings).Trim()))
            .Where(header => !string.IsNullOrWhiteSpace(header.Value))
            .OrderBy(header => header.Key).ToList();
        if (headers.Count == 0 || headers.Any(header => header.Key <= 0 || header.Key == int.MaxValue) ||
            headers.Select(header => header.Value).Distinct(StringComparer.OrdinalIgnoreCase).Count() != headers.Count)
            throw new InvalidDataException("The worksheet must have unique, nonempty column headings with valid cell references.");
        List<string> columns = headers.Select(header => header.Value).ToList();

        List<Dictionary<string, string>> rows = worksheetXml
            .Descendants(SpreadsheetNs + "sheetData")
            .Elements(SpreadsheetNs + "row")
            .Skip(1)
            .Select(row => ReadRow(row, headers, sharedStrings))
            .Where(row => row.Values.Any(value => !string.IsNullOrWhiteSpace(value)))
            .ToList();

        return new RecipientFileData(columns, rows);
    }

    private static Dictionary<string, string> ReadRow(XElement row, List<KeyValuePair<int, string>> headers, List<string> sharedStrings)
    {
        Dictionary<int, string> valuesByColumnIndex = row
            .Elements(SpreadsheetNs + "c")
            .ToDictionary(
                cell => GetColumnIndex((string?)cell.Attribute("r")),
                cell => ReadCellText(cell, sharedStrings));

        Dictionary<string, string> recipientRow = new(StringComparer.OrdinalIgnoreCase);
        for (int index = 0; index < headers.Count; index++)
        {
            recipientRow[headers[index].Value] = valuesByColumnIndex.TryGetValue(headers[index].Key, out string? value)
                ? value.Trim()
                : string.Empty;
        }

        return recipientRow;
    }

    private static List<string> ReadSharedStrings(ZipArchive archive)
    {
        ZipArchiveEntry? sharedStringsEntry = archive.GetEntry("xl/sharedStrings.xml");
        if (sharedStringsEntry is null)
        {
            return [];
        }

        XDocument sharedStrings = ReadXml(sharedStringsEntry);
        return sharedStrings
            .Descendants(SpreadsheetNs + "si")
            .Select(ReadSharedString)
            .ToList();
    }

    private static string ReadSharedString(XElement sharedString)
    {
        IEnumerable<XElement> textParts = sharedString.Descendants(SpreadsheetNs + "t");
        return string.Concat(textParts.Select(textPart => textPart.Value));
    }

    private static string ReadCellText(XElement cell, List<string> sharedStrings)
    {
        string cellType = (string?)cell.Attribute("t") ?? string.Empty;

        if (cellType == "s")
        {
            string rawIndex = cell.Element(SpreadsheetNs + "v")?.Value ?? string.Empty;
            return int.TryParse(rawIndex, out int sharedStringIndex) &&
                   sharedStringIndex >= 0 &&
                   sharedStringIndex < sharedStrings.Count
                ? sharedStrings[sharedStringIndex]
                : string.Empty;
        }

        if (cellType == "inlineStr")
        {
            return string.Concat(cell.Descendants(SpreadsheetNs + "t").Select(part => part.Value));
        }

        return cell.Element(SpreadsheetNs + "v")?.Value ?? string.Empty;
    }

    private static int GetColumnIndex(string? cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference))
        {
            return int.MaxValue;
        }

        int columnIndex = 0;
        foreach (char character in cellReference.Where(char.IsLetter))
        {
            columnIndex *= 26;
            columnIndex += char.ToUpperInvariant(character) - 'A' + 1;
        }

        return columnIndex;
    }

    private static string NormalizeWorkbookPartPath(string target)
    {
        string normalized = target.Replace('\\', '/').TrimStart('/');
        if (!normalized.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
        {
            normalized = $"xl/{normalized}";
        }

        return normalized;
    }

    private static XDocument ReadXml(ZipArchiveEntry entry)
    {
        using Stream stream = entry.Open();
        using StreamReader reader = new(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return XDocument.Load(reader);
    }
}

public sealed record ExcelWorksheetInfo(
    string Name,
    string WorksheetPath);
