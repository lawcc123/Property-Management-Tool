using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PropertyManager.App.Envelopes;

internal sealed record MailingAddress
{
    [JsonPropertyName("id")] public string Id { get; init; } = Guid.NewGuid().ToString("N");
    [JsonPropertyName("nickname")] public string Nickname { get; init; } = "";
    [JsonPropertyName("company")] public string Company { get; init; } = "";
    [JsonPropertyName("address")] public string Street { get; init; } = "";
    [JsonPropertyName("city")] public string City { get; init; } = "";
    [JsonPropertyName("province")] public string Province { get; init; } = "";
    [JsonPropertyName("postal")] public string Postal { get; init; } = "";
    [JsonPropertyName("logo_png")] public string? LogoPng { get; init; }
    [JsonPropertyName("logo_width_mm")] public double LogoWidthMm { get; init; }=22;
    [JsonIgnore] public string DisplayName => string.IsNullOrWhiteSpace(Nickname) ? Company : Nickname;
    [JsonIgnore] public string CityLine => string.Join("  ", new[]{string.Join(", ",new[]{City,Province}.Where(s=>!string.IsNullOrWhiteSpace(s))),Postal}.Where(s=>!string.IsNullOrWhiteSpace(s)));
    [JsonIgnore] public string Text => string.Join("\n",new[]{Company,Street,CityLine}.Where(s=>!string.IsNullOrWhiteSpace(s)));
    internal void Validate()
    {
        if(string.IsNullOrWhiteSpace(Company)||string.IsNullOrWhiteSpace(Street)||string.IsNullOrWhiteSpace(City))
            throw new InvalidOperationException("Each address needs a company/name, street address, and city.");
        if(!double.IsFinite(LogoWidthMm)||LogoWidthMm<10||LogoWidthMm>40)throw new InvalidOperationException("Use a logo width from 10 to 40 mm.");
        if(LogoPng is not null)ReturnAddressLogo.Decode(LogoPng);
        if(Text.Length>2000)throw new InvalidOperationException("The address is too long. Shorten it before printing.");
    }
}
internal sealed class AddressBook
{
    public List<MailingAddress> Contractors {get;set;}=[];
    public List<MailingAddress> Offices {get;set;}=[];
    internal void Validate()
    {
        foreach(var list in new[]{Contractors,Offices})
            if(list.Any(a=>string.IsNullOrWhiteSpace(a.Id))||list.Select(a=>a.Id).Distinct(StringComparer.Ordinal).Count()!=list.Count)
                throw new InvalidOperationException("Address IDs must be present and unique. The original files have not been changed.");
    }
}
internal sealed class RecipientRow : INotifyPropertyChanged
{
    internal RecipientRow(MailingAddress address){Address=address;}
    public MailingAddress Address {get;}
    public string Company=>Address.Company;
    public string City=>Address.City;
    public string Postal=>Address.Postal;
    private bool selected;
    public bool Selected {get=>selected;set{if(selected==value)return;selected=value;PropertyChanged?.Invoke(this,new(nameof(Selected)));}}
    public event PropertyChangedEventHandler? PropertyChanged;
}
internal sealed record EnvelopePreset(string Key,string Label,double WidthMm,double HeightMm)
{
    internal static readonly EnvelopePreset[] All =
    [new("10","#10 envelope - 241 x 104 mm",241,104),new("DL","DL — 220 × 110 mm",220,110),
     new("C5","C5 — 229 × 162 mm",229,162),new("C6","C6 — 162 × 114 mm",162,114),
     new("Monarch","Monarch — 190.5 × 98.4 mm",190.5,98.425),new("Custom","Custom dimensions",241.3,104.775)];
}
internal sealed class EnvelopeOptions
{
    public int LayoutVersion {get;set;}=1;
    public bool IncludeReturnLogo {get;set;}
    [JsonIgnore] internal int ExpectedLayoutVersion=>IncludeReturnLogo?4:IsLabel?(ReturnLabelCount>0?3:2):1;
    public LabelLayout? LabelSheet {get;set;}
    public int FirstLabelSlot {get;set;}=1;
    public List<int>? FirstPageSlots {get;set;}
    [JsonIgnore] public bool IsLabel=>LabelSheet is not null;
    public string Preset {get;set;}="10";
    public double WidthMm {get;set;}=241;
    public double HeightMm {get;set;}=104;
    public double FontPt {get;set;}=14;
    public double ReturnFontPt {get;set;}=9;
    public double OffsetXmm {get;set;}
    public double OffsetYmm {get;set;}
    public bool DefaultToBypass {get;set;}=true;
    public bool WordEnvelopeCompatibility {get;set;}
    [JsonIgnore] internal bool UsesWordEnvelopeCompatibility=>!IsLabel&&(WordEnvelopeCompatibility||Preset=="10"&&Math.Abs(WidthMm-241)<0.01&&Math.Abs(HeightMm-104)<0.01);
    public bool Rotate180 {get;set;}
    public int Copies {get;set;}=1;
    public int ReturnLabelCount {get;set;}
    public string Printer {get;set;}="";
    public int? TrayId {get;set;}
    public string TrayName {get;set;}="";
    public string ReturnAddressId {get;set;}="";
    internal EnvelopeOptions Copy()=>JsonSerializer.Deserialize<EnvelopeOptions>(JsonSerializer.Serialize(this))!;
    internal EnvelopeOptions CurrentEnvelopeSettings()
    {
        var current=Copy();
        if(!IsLabel&&(Preset=="10"||WordEnvelopeCompatibility&&Preset=="Custom"&&Math.Abs(WidthMm-241)<0.01&&Math.Abs(HeightMm-104)<0.01))
        {current.Preset="10";current.WidthMm=241;current.HeightMm=104;current.WordEnvelopeCompatibility=false;}
        return current;
    }
    internal void Validate()
    {
        if(LayoutVersion!=ExpectedLayoutVersion)throw new InvalidOperationException("This saved layout version is not supported by this app.");
        if(!double.IsFinite(WidthMm)||!double.IsFinite(HeightMm)||WidthMm<50||WidthMm>450||HeightMm<50||HeightMm>450)
            throw new InvalidOperationException("Envelope width and height must be between 50 and 450 mm.");
        if(!double.IsFinite(FontPt)||!double.IsFinite(ReturnFontPt)||FontPt<6||FontPt>36||ReturnFontPt<6||ReturnFontPt>18)
            throw new InvalidOperationException("Use recipient text from 6–36 pt and return-address text from 6–18 pt.");
        if(!double.IsFinite(OffsetXmm)||!double.IsFinite(OffsetYmm)||Math.Abs(OffsetXmm)>40||Math.Abs(OffsetYmm)>40)
            throw new InvalidOperationException("Alignment offsets must be between −40 and 40 mm.");
        if(LabelSheet is {} sheet)
        {
            sheet.Validate();
            if(Math.Abs(WidthMm-215.9)>0.01||Math.Abs(HeightMm-279.4)>0.01)throw new InvalidOperationException("These label templates require US Letter sheets (215.9 x 279.4 mm).");
            if(FirstLabelSlot<1||FirstLabelSlot>sheet.Capacity)throw new InvalidOperationException("Choose a valid starting label slot.");
            if(FirstPageSlots is {} slots && (slots.Count==0||slots.Distinct().Count()!=slots.Count||slots.Any(n=>n<1||n>sheet.Capacity)))throw new InvalidOperationException("Choose at least one valid, unique label slot on the first sheet.");
        }
        if(ReturnLabelCount<0||ReturnLabelCount>10000||!IsLabel&&ReturnLabelCount!=0)throw new InvalidOperationException("Choose up to 10,000 return-address labels in label mode.");
        if(Copies<1||Copies>99)throw new InvalidOperationException("Choose 1–99 copies per recipient.");
    }
}
internal sealed record EnvelopeSnapshot(List<MailingAddress> Recipients,MailingAddress? ReturnAddress,EnvelopeOptions Options)
{
    [JsonIgnore] public int ContractorItems=>checked(Recipients.Count*Options.Copies);
    [JsonIgnore] public int Items=>checked(ContractorItems+Options.ReturnLabelCount);
    [JsonIgnore] public int Pages=>Options.LabelSheet is {} sheet
        ? 1+(int)Math.Ceiling(Math.Max(0,Items-(Options.FirstPageSlots?.Count??sheet.Capacity-Options.FirstLabelSlot+1))/(double)sheet.Capacity) : Items;
    internal EnvelopeSnapshot Copy()=>new(Recipients.Select(a=>a with{}).ToList(),ReturnAddress is null?null:ReturnAddress with{},Options.Copy());
    internal void Validate()
    {
        Options.Validate();
        if(Options.IncludeReturnLogo&&ReturnAddress?.LogoPng is null)throw new InvalidOperationException("Choose a return address with a saved logo, or turn off logo printing.");
        if(Recipients.Count==0&&Options.ReturnLabelCount==0)throw new InvalidOperationException("Select at least one recipient or enable return-address labels.");
        if(Options.ReturnLabelCount>0&&ReturnAddress is null)throw new InvalidOperationException("Choose a return address for the return-address labels.");
        if(Items>10000)throw new InvalidOperationException("Split this request into batches of at most 10,000 envelopes or labels.");
        foreach(var a in Recipients)a.Validate();ReturnAddress?.Validate();
    }
}
internal sealed record HistoryEvent(DateTimeOffset WhenUtc,string Status,string Details);
internal sealed class EnvelopeHistoryEntry
{
    public string Id {get;set;}=Guid.NewGuid().ToString("N");
    public string? ParentId {get;set;}
    public DateTimeOffset RequestedUtc {get;set;}=DateTimeOffset.UtcNow;
    public string WindowsUser {get;set;}=Environment.UserName;
    public string Kind {get;set;}="Print";
    public EnvelopeSnapshot Snapshot {get;set;}=new([],null,new());
    public int? JobId {get;set;}
    public string DocumentName {get;set;}="";
    public List<HistoryEvent> Events {get;set;}=[];
    [JsonIgnore] public string Status=>Events.LastOrDefault()?.Status??"Requested";
    [JsonIgnore] public string RequestedLocal=>TimeZoneInfo.ConvertTime(RequestedUtc,TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")).ToString("yyyy-MM-dd HH:mm:ss zzz");
    [JsonIgnore] public string RecipientSummary=>string.Join("; ",new[]{string.Join(", ",Snapshot.Recipients.Take(3).Select(a=>a.Company))+(Snapshot.Recipients.Count>3?" ...":""),Snapshot.Options.ReturnLabelCount>0?$"Return labels: {Snapshot.ReturnAddress?.Company} ({Snapshot.Options.ReturnLabelCount})":""}.Where(value=>value.Length>0));
    [JsonIgnore] public string EnvelopeSize=>Snapshot.Options.LabelSheet?.Label??$"{Snapshot.Options.WidthMm:0.#} \u00D7 {Snapshot.Options.HeightMm:0.#} mm";
    [JsonIgnore] public int Pages=>Snapshot.Pages;
}
