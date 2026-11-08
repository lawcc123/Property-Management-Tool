using System.IO;
using System.ComponentModel;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;

namespace PropertyManager.App.EmailLists;

internal enum EmailListMode { Risers, Floors, Residents, OnsiteOwners, AllOwners, Retail, Units }
internal enum EmailListAudience { Residents, OnsiteOwners, AllOwners, AllOccupants }

internal sealed class GeneratedRecipient : INotifyPropertyChanged
{
    public string Unit { get; set; }
    public string Floor { get; set; }
    public string Riser { get; set; }
    public string LastName { get; set; }
    public string FirstName { get; set; }
    public string Names { get; set; }
    public string UserType { get; set; }
    public AlternateMailingAddress Mailing {get;set;}=new();
    public string MailingName=>Mailing.Name;
    public string MailingAddress=>Mailing.Text;
    public string MailingStatus=>Mailing.Status;
    public bool NeedsMailingReview=>MailingStatus.Contains("review",StringComparison.OrdinalIgnoreCase);
    private string _email;
    public string Email { get => _email; set { _email = value; Changed(nameof(Email)); } }
    private string _status;
    public string Status { get => _status; set { _status = value; Changed(nameof(Status)); } }
    public bool IsValid => GeneratedEmailListService.EmailParts(Email) is { Length: > 0 } parts && parts.All(GeneratedEmailListService.IsValidEmail);
    public event PropertyChangedEventHandler? PropertyChanged;
    private void Changed(string name) => PropertyChanged?.Invoke(this, new(name));

    public GeneratedRecipient(string unit, string floor, string riser, string lastName,
        string firstName, string userType, string email, string status)
    {
        Unit = unit; Floor = floor; Riser = riser; LastName = lastName;
        FirstName = firstName; UserType = userType; _email = email; _status = status;
        Names = string.Join(" ", new[] { firstName, lastName }.Where(s => s.Length > 0));
    }
}

internal static class GeneratedEmailListService
{
    private static readonly string[] RequiredColumns =
        ["Building Name", "Unit ID", "Floor", "First Name", "Last Name", "Emails", "User Type", "User Active?", "Unit Active?"];
    public static readonly string[] ExportColumns =
        ["Unit", "Floor", "Riser", "Last Name", "First Name", "User Type", "Email", "Status"];
    public static readonly string[] GroupedExportColumns =
        ["Unit", "Floor", "Riser", "Names", "User Type", "Email", "Status"];

    public static void Validate(RecipientFileData data)
    {
        string[] missing = RequiredColumns.Where(c => !data.Columns.Contains(c, StringComparer.OrdinalIgnoreCase)).ToArray();
        if (missing.Length > 0)
            throw new InvalidDataException("This is not a supported BuildingLink occupant export. Missing columns: " + string.Join(", ", missing));
        if (data.Rows.Count == 0)
            throw new InvalidDataException("The occupant export contains no data rows.");
        if (data.Rows.Select(r => Value(r, "Building Name")).Where(v => v.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase).Skip(1).Any())
            throw new InvalidDataException("Choose occupant data for a single corporation. This file contains multiple buildings.");
    }

    internal static string Value(Dictionary<string, string> row, string column) =>
        row.TryGetValue(column, out string? value) ? value.Trim() : string.Empty;

    public static string Riser(Dictionary<string, string> row, bool useLine)
    {
        if (IsRetail(row)) return string.Empty;
        if (useLine) return Value(row, "Line");
        string unit = Value(row, "Unit ID");
        // Only numeric units and PH units have an unambiguous numeric suffix.
        Match match = Regex.Match(unit, @"^(?:PH\s*)?(\d+)$", RegexOptions.IgnoreCase);
        if (!match.Success) return string.Empty;
        string digits = match.Groups[1].Value;
        return digits.Length >= 2 ? digits[^2..] : digits.PadLeft(2, '0');
    }

    private static bool IsRetail(Dictionary<string, string> row) =>
        Value(row, "Unit ID").StartsWith("Retail", StringComparison.OrdinalIgnoreCase);
    private static bool IsOwner(Dictionary<string, string> row) =>
        Value(row, "User Type").Equals("Owner", StringComparison.OrdinalIgnoreCase);
    private static bool IsAnyOwner(Dictionary<string, string> row) => IsOwner(row) ||
        new[]{"Offsite Owner","Off-site Owner"}.Contains(Value(row, "User Type"),StringComparer.OrdinalIgnoreCase);
    private static bool IsResident(Dictionary<string, string> row)
    {
        string type = Value(row, "User Type");
        return !IsRetail(row) && (IsOwner(row) ||
            new[] { "Resident", "Renter", "Tenant" }.Contains(type, StringComparer.OrdinalIgnoreCase) ||
            type.StartsWith("Family Member", StringComparison.OrdinalIgnoreCase));
    }

    public static List<GeneratedRecipient> Generate(RecipientFileData data, EmailListMode mode,
        EmailListAudience audience, IEnumerable<string> selectedValues, bool useLine, bool activeOnly, bool deduplicate,
        bool groupByUnit = false, bool includeMailing = false, RecipientFileData? mailingData = null, string? defaultCorporation = null, bool ownerMailingRules = false)
    {
        Validate(data);
        if(ownerMailingRules){mode=EmailListMode.AllOwners;includeMailing=true;activeOnly=true;groupByUnit=true;deduplicate=true;if(mailingData is null)throw new InvalidDataException("Load both Occupant Data and Unit Data for owner mailing lists.");}
        if(includeMailing&&defaultCorporation is not null&&mailingData is null&&!AlternateMailingAddress.HasColumns(data))throw new InvalidDataException("Load Unit Data to verify alternate addresses before applying default property addresses.");
        var mailingByUnit=includeMailing&&mailingData is not null?AlternateMailingAddress.UnitMap(mailingData):null;
        HashSet<string> selected = new(selectedValues, StringComparer.OrdinalIgnoreCase);
        List<GeneratedRecipient> result = [];
        foreach (Dictionary<string, string> row in data.Rows)
        {
            if (activeOnly && (!Value(row, "User Active?").Equals("Y", StringComparison.OrdinalIgnoreCase) ||
                               !Value(row, "Unit Active?").Equals("Y", StringComparison.OrdinalIgnoreCase))) continue;
            bool matches = mode switch
            {
                EmailListMode.Residents => IsResident(row),
                EmailListMode.OnsiteOwners => !IsRetail(row) && IsOwner(row),
                EmailListMode.AllOwners => IsAnyOwner(row),
                EmailListMode.Retail => IsRetail(row),
                _ => selected.Contains(mode == EmailListMode.Units ? Value(row, "Unit ID") : mode == EmailListMode.Floors ? Value(row, "Floor") : Riser(row, useLine)) &&
                     (audience switch
                     {
                         EmailListAudience.Residents => IsResident(row),
                         EmailListAudience.OnsiteOwners => !IsRetail(row) && IsOwner(row),
                         EmailListAudience.AllOwners => IsAnyOwner(row),
                         _ => true
                     })
            };
            if (!matches) continue;
            var mailing=includeMailing?AlternateMailingAddress.Read(row,mailingByUnit):new();
            if(includeMailing&&defaultCorporation is not null)mailing=OwnerMailingDefault.Apply(mailing,defaultCorporation,Value(row,"Unit ID"),string.Join(" ",new[]{Value(row,"First Name"),Value(row,"Last Name")}.Where(v=>v.Length>0)));
            // One row per contact keeps document recipients even when they share an email.
            result.Add(new(Value(row, "Unit ID"), Value(row, "Floor"), Riser(row, useLine),
                Value(row, "Last Name"), Value(row, "First Name"), Value(row, "User Type"), Value(row, "Emails"), "")
            {Mailing=mailing});
        }
        if (groupByUnit) result = ownerMailingRules?GroupOwnersByUnit(result):GroupByUnit(result);
        // Grouped notices keep shared addresses across units, but deduplicate within each unit.
        if(groupByUnit){foreach(var row in result)RefreshEmails([row],true);}else RefreshEmails(result, deduplicate);
        return result;
    }

    public static string[] EmailParts(string email) => email.Split([';', ',', '\r', '\n'],
        StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    public static bool IsValidEmail(string email) => MailAddress.TryCreate(email, out MailAddress? address) &&
        address.Address.Equals(email, StringComparison.OrdinalIgnoreCase);

    public static List<GeneratedRecipient> GroupByUnit(IEnumerable<GeneratedRecipient> contacts) => contacts
        .GroupBy(r => r.Unit.ToUpperInvariant()+"\u001F"+r.Mailing.Key).Select(group =>
        {
            GeneratedRecipient first = group.First();
            return new GeneratedRecipient(first.Unit, first.Floor, first.Riser, "", "",
                string.Join(" & ", group.Select(r => r.UserType).Distinct(StringComparer.OrdinalIgnoreCase)),
                string.Join("; ", group.SelectMany(r => EmailParts(r.Email)).Distinct(StringComparer.OrdinalIgnoreCase)), "")
            {
                Mailing=first.Mailing with {Name=string.Join(" & ",group.Select(r=>r.Mailing.Name).Where(n=>n.Length>0).Distinct(StringComparer.OrdinalIgnoreCase)),Status=string.Join("; ",group.Select(r=>r.Mailing.Status).Where(n=>n.Length>0).Distinct())},
                Names = string.Join(" & ", group.Select(r => r.Names).Where(n => n.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase))
            };
        }).ToList();

    private static List<GeneratedRecipient> GroupOwnersByUnit(IEnumerable<GeneratedRecipient> contacts)
    {
        return contacts.GroupBy(r=>AlternateMailingAddress.UnitKey(r.Unit),StringComparer.OrdinalIgnoreCase).Select(unit=>
        {
            var first=unit.First();var addresses=unit.Select(r=>r.Mailing).GroupBy(a=>a.Key).Select(g=>g.First()).ToList();
            var address=addresses.Count==1?first.Mailing:new AlternateMailingAddress
            {
                Line1=string.Join("\n\n",addresses.Select(a=>a.Text).Where(t=>t.Length>0)),
                Status="Review: different mailing addresses for this unit; choose one before mailing"
            };
            address=address with {Name=string.Join(" & ",unit.Select(r=>r.Mailing.Name).Where(n=>n.Length>0).Distinct(StringComparer.OrdinalIgnoreCase))};
            return new GeneratedRecipient(first.Unit,first.Floor,first.Riser,"","",string.Join(" & ",unit.Select(r=>r.UserType).Distinct(StringComparer.OrdinalIgnoreCase)),
                string.Join("; ",unit.SelectMany(r=>EmailParts(r.Email)).Distinct(StringComparer.OrdinalIgnoreCase)),"")
            {Names=string.Join(" & ",unit.Select(r=>r.Names).Where(n=>n.Length>0).Distinct(StringComparer.OrdinalIgnoreCase)),Mailing=address};
        }).ToList();
    }
    public static void RefreshEmails(IEnumerable<GeneratedRecipient> recipients, bool deduplicate)
    {
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
        foreach (GeneratedRecipient row in recipients)
        {
            string[] parts = EmailParts(row.Email);
            // A row that cannot be sent must not consume an address needed by a valid row.
            if (parts.Any(address => !IsValidEmail(address)))
            {
                row.Email = string.Join("; ", parts);
                row.Status = "Invalid email";
                continue;
            }
            List<string> kept = [];
            int removed = 0;
            bool duplicate = false;
            foreach (string address in parts)
            {
                if (IsValidEmail(address) && !seen.Add(address))
                {
                    duplicate = true;
                    if (deduplicate) { removed++; continue; }
                }
                kept.Add(address);
            }
            row.Email = string.Join("; ", kept);
            row.Status = kept.Count == 0 ? (removed > 0 || row.Status == "Duplicate removed - print only"
                ? "Duplicate removed - print only" : "Missing email")
                : kept.Any(e => !IsValidEmail(e)) ? "Invalid email"
                : removed > 0 ? "Ready (duplicates removed)" : duplicate ? "Duplicate email" : "Ready";
        }
    }

    public static RecipientFileData ToRecipientData(IEnumerable<GeneratedRecipient> recipients, bool grouped = false, bool includeMailing = false)
    {
        string[] columns = grouped ? GroupedExportColumns : ExportColumns;
        if(includeMailing)columns=[..columns,..AlternateMailingAddress.ExportColumns];
        return new([.. columns], recipients.Select(r =>
        {
            string[] values = grouped ? [r.Unit, r.Floor, r.Riser, r.Names, r.UserType, r.Email, r.Status] :
                [r.Unit, r.Floor, r.Riser, r.LastName, r.FirstName, r.UserType, r.Email, r.Status];
            if(includeMailing)values=[..values,..r.Mailing.ExportValues];
            return columns.Zip(values).ToDictionary(p => p.First, p => p.Second, StringComparer.OrdinalIgnoreCase);
        }).ToList());
    }

    public static void Export(string path, IEnumerable<GeneratedRecipient> recipients, bool grouped = false, bool includeMailing = false)
    {
        RecipientFileData data = ToRecipientData(recipients, grouped, includeMailing);
        if(includeMailing)
        {
            var omitted=new HashSet<string>(StringComparer.OrdinalIgnoreCase){"Mailing Name","Floor","Riser","Status","Mailing Address Status"};
            var columns=data.Columns.Where(c=>!omitted.Contains(c)).ToList();
            data=new(columns,data.Rows.Select(row=>columns.ToDictionary(c=>c,c=>row[c],StringComparer.OrdinalIgnoreCase)).ToList());
        }
        if (Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            ExcelRecipientExporter.Export(path, data);
            return;
        }
        using StreamWriter writer = new(path, false, new UTF8Encoding(true));
        static string Escape(string value)
        {
            // Keep spreadsheet applications from treating imported text as a formula.
            if (value.TrimStart().StartsWith('=') || value.TrimStart().StartsWith('+') ||
                value.TrimStart().StartsWith('-') || value.TrimStart().StartsWith('@')) value = "'" + value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
        writer.WriteLine(string.Join(",", data.Columns.Select(Escape)));
        foreach (var row in data.Rows) writer.WriteLine(string.Join(",", data.Columns.Select(c => Escape(row[c]))));
    }
}
