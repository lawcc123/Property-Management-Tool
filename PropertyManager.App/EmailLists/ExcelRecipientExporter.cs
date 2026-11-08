using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace PropertyManager.App.EmailLists;

internal static class ExcelRecipientExporter
{
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/package/2006/content-types";

    public static void Export(string path, RecipientFileData data)
    {
        if (data.Columns.Count == 0 || data.Rows.Count == 0)
            throw new InvalidDataException("There are no recipient rows to export.");
        string range = $"A1:{ColumnName(data.Columns.Count)}{data.Rows.Count + 1}";
        double[] widths = data.Columns.Select(c => Math.Clamp(
            data.Rows.Select(r => Width(r[c])).Append(Width(c)).Max() + 3, 12, c == "Email" ? 70 : 60)).ToArray();
        using FileStream output = new(path, FileMode.Create, FileAccess.Write);
        using ZipArchive zip = new(output, ZipArchiveMode.Create);
        Write(zip, "[Content_Types].xml", new XElement(C + "Types",
            new XElement(C + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
            new XElement(C + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
            Content("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
            Content("/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
            Content("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
            Content("/xl/tables/table1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")));
        Write(zip, "_rels/.rels", Relationships(Rel("rId1", "officeDocument", "xl/workbook.xml")));
        Write(zip, "xl/workbook.xml", new XElement(S + "workbook", new XAttribute(XNamespace.Xmlns + "r", R),
            new XElement(S + "sheets", new XElement(S + "sheet", new XAttribute("name", "Generated recipients"),
                new XAttribute("sheetId", 1), new XAttribute(R + "id", "rId1")))));
        Write(zip, "xl/_rels/workbook.xml.rels", Relationships(Rel("rId1", "worksheet", "worksheets/sheet1.xml"), Rel("rId2", "styles", "styles.xml")));
        Write(zip, "xl/styles.xml", new XElement(S + "styleSheet",
            new XElement(S + "fonts", new XAttribute("count", 1), new XElement(S + "font", new XElement(S + "sz", new XAttribute("val", 11)), new XElement(S + "name", new XAttribute("val", "Calibri")))),
            new XElement(S + "fills", new XAttribute("count", 2), new XElement(S + "fill", new XElement(S + "patternFill", new XAttribute("patternType", "none"))), new XElement(S + "fill", new XElement(S + "patternFill", new XAttribute("patternType", "gray125")))),
            new XElement(S + "borders", new XAttribute("count", 1), new XElement(S + "border")),
            new XElement(S + "cellStyleXfs", new XAttribute("count", 1), Xf(false)),
            new XElement(S + "cellXfs", new XAttribute("count", 1), Xf(true)),
            new XElement(S + "cellStyles", new XAttribute("count", 1), new XElement(S + "cellStyle", new XAttribute("name", "Normal"), new XAttribute("xfId", 0), new XAttribute("builtinId", 0))),
            new XElement(S + "tableStyles", new XAttribute("count", 0), new XAttribute("defaultTableStyle", "TableStyleMedium2"), new XAttribute("defaultPivotStyle", "PivotStyleLight16"))));
        List<XElement> rows = [Row(1, data.Columns, widths)];
        for (int i = 0; i < data.Rows.Count; i++)
            rows.Add(Row(i + 2, data.Columns.Select(c => data.Rows[i][c]).ToArray(), widths));
        Write(zip, "xl/worksheets/sheet1.xml", new XElement(S + "worksheet", new XAttribute(XNamespace.Xmlns + "r", R),
            new XElement(S + "dimension", new XAttribute("ref", range)),
            new XElement(S + "sheetViews", new XElement(S + "sheetView", new XAttribute("workbookViewId", 0),
                new XElement(S + "pane", new XAttribute("ySplit", 1), new XAttribute("topLeftCell", "A2"), new XAttribute("activePane", "bottomLeft"), new XAttribute("state", "frozen")))),
            new XElement(S + "sheetFormatPr", new XAttribute("defaultRowHeight", 18)),
            new XElement(S + "cols", widths.Select((w, i) => new XElement(S + "col", new XAttribute("min", i + 1), new XAttribute("max", i + 1), new XAttribute("width", w.ToString(CultureInfo.InvariantCulture)), new XAttribute("customWidth", 1)))),
            new XElement(S + "sheetData", rows),
            new XElement(S + "tableParts", new XAttribute("count", 1), new XElement(S + "tablePart", new XAttribute(R + "id", "rId1")))));
        Write(zip, "xl/worksheets/_rels/sheet1.xml.rels", Relationships(Rel("rId1", "table", "../tables/table1.xml")));
        Write(zip, "xl/tables/table1.xml", new XElement(S + "table", new XAttribute("id", 1), new XAttribute("name", "GeneratedRecipients"),
            new XAttribute("displayName", "GeneratedRecipients"), new XAttribute("ref", range), new XAttribute("totalsRowShown", 0),
            new XElement(S + "autoFilter", new XAttribute("ref", range)),
            new XElement(S + "tableColumns", new XAttribute("count", data.Columns.Count), data.Columns.Select((c, i) => new XElement(S + "tableColumn", new XAttribute("id", i + 1), new XAttribute("name", c)))),
            new XElement(S + "tableStyleInfo", new XAttribute("name", "TableStyleMedium2"), new XAttribute("showFirstColumn", 0), new XAttribute("showLastColumn", 0), new XAttribute("showRowStripes", 1), new XAttribute("showColumnStripes", 0))));
    }

    private static XElement Xf(bool wrap) => new(S + "xf", new XAttribute("numFmtId", wrap ? 49 : 0), new XAttribute("fontId", 0),
        new XAttribute("fillId", 0), new XAttribute("borderId", 0), new XAttribute("xfId", 0),
        wrap ? new XAttribute("applyAlignment", 1) : null, wrap ? new XAttribute("applyNumberFormat", 1) : null,
        wrap ? new XElement(S + "alignment", new XAttribute("vertical", "top"), new XAttribute("wrapText", 1)) : null);
    private static double Width(string text) => text.Split('\n').Select(line => line.Sum(ch => ch >= 0x2e80 ? 2.0 : 1.0)).DefaultIfEmpty(0).Max();
    private static XElement Row(int index, IList<string> values, double[] widths)
    {
        double lines = values.Select((v, i) => v.Split('\n').Sum(line => Math.Max(1, Math.Ceiling(Width(line) / Math.Max(1, widths[i] - 3))))).Max();
        return new(S + "row", new XAttribute("r", index), new XAttribute("ht", Math.Min(409, Math.Max(20, lines * 16 + 4))), new XAttribute("customHeight", 1),
            values.Select((v, i) => new XElement(S + "c", new XAttribute("r", $"{ColumnName(i + 1)}{index}"), new XAttribute("t", "inlineStr"), new XAttribute("s", 0),
                new XElement(S + "is", new XElement(S + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), new string(v.Where(XmlConvert.IsXmlChar).ToArray()))))));
    }
    private static string ColumnName(int index)
    {
        string result = "";
        while (index > 0) { index--; result = (char)('A' + index % 26) + result; index /= 26; }
        return result;
    }
    private static XElement Content(string part, string type) => new(C + "Override", new XAttribute("PartName", part), new XAttribute("ContentType", type));
    private static XElement Relationships(params XElement[] entries) => new(P + "Relationships", entries);
    private static XElement Rel(string id, string type, string target) => new(P + "Relationship", new XAttribute("Id", id), new XAttribute("Type", $"{R}/{type}"), new XAttribute("Target", target));
    private static void Write(ZipArchive zip, string path, XElement root)
    {
        using Stream stream = zip.CreateEntry(path).Open();
        new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root).Save(stream);
    }
}
