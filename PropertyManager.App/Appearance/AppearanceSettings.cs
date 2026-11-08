using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Media;

namespace PropertyManager.App.Appearance;

internal sealed record WorkspaceItem(string Id,string Name);
internal sealed record ThemePreset(string Name,string Accent,string Sidebar);
internal sealed class AppearanceSettings
{
    internal static readonly WorkspaceItem[] Workspaces=[new("registry","Registry Sync"),new("email","Generate Email List"),new("owners","Owner Mailing List"),new("buildinglink","BuildingLink Data"),new("printing","Printing contractor's address"),new("outlook","Outlook Contact Sync"),new("reminders","Reservations & Reminders"),new("personalized","Personalized Email")];
    internal static readonly ThemePreset[] Presets=[new("Teal (original)","#147D70","#142D3B"),new("Ocean blue","#2463A4","#152C44"),new("Plum","#795399","#302139"),new("Burgundy","#A33D60","#351F2B"),new("Forest","#397342","#1D3027"),new("Slate","#526575","#242D36"),new("Muted yellow (shared)","#C9AB48","#332D27")];
    public string Accent {get;set;}="#147D70";
    public string Sidebar {get;set;}="#142D3B";
    public List<string> WorkspaceOrder {get;set;}=Workspaces.Select(w=>w.Id).ToList();
    internal static string SettingsPath=>Path.Combine(Services.AppSettings.DataDirectory,"appearance-settings.json");
    internal AppearanceSettings Copy()=>new(){Accent=Accent,Sidebar=Sidebar,WorkspaceOrder=WorkspaceOrder.ToList()};
    internal void Normalize()
    {
        WorkspaceOrder=(WorkspaceOrder??[]).Where(id=>Workspaces.Any(w=>w.Id==id)).Distinct().Concat(Workspaces.Select(w=>w.Id)).Distinct().ToList();
    }
    internal void Validate()
    {
        foreach(string? color in new[]{Accent,Sidebar})if(color is null||!Regex.IsMatch(color,"^#[0-9A-Fa-f]{6}$"))throw new InvalidDataException("Enter a colour as #RRGGBB, for example #2463A4.");
        if(Contrast(Parse(Sidebar),Colors.White)<7)throw new InvalidDataException("Choose a darker sidebar colour so its white labels remain easy to read.");
        Accent=Accent.ToUpperInvariant();Sidebar=Sidebar.ToUpperInvariant();Normalize();
    }
    internal static AppearanceSettings Load(string? path=null)
    {
        var settings=Settings.SettingsDatabase.Read(path??SettingsPath) is {} json?JsonSerializer.Deserialize<AppearanceSettings>(json)??throw new InvalidDataException("Appearance settings are empty."):new();settings.Validate();return settings;
    }
    internal void Save(string? path=null)
    {
        Validate();path??=SettingsPath;Settings.SettingsDatabase.Write(path,JsonSerializer.Serialize(this,new JsonSerializerOptions{WriteIndented=true}));
    }
    internal static Color Parse(string color)=>(Color)ColorConverter.ConvertFromString(color);
    internal static Color Mix(Color a,Color b,double amount)=>Color.FromRgb((byte)Math.Round(a.R*(1-amount)+b.R*amount),(byte)Math.Round(a.G*(1-amount)+b.G*amount),(byte)Math.Round(a.B*(1-amount)+b.B*amount));
    internal static double Contrast(Color a,Color b)
    {
        double L(Color c){double V(byte value){double v=value/255.0;return v<=0.04045?v/12.92:Math.Pow((v+0.055)/1.055,2.4);}return 0.2126*V(c.R)+0.7152*V(c.G)+0.0722*V(c.B);}
        double x=L(a),y=L(b);return (Math.Max(x,y)+0.05)/(Math.Min(x,y)+0.05);
    }
    internal void Apply()
    {
        Validate();var resources=Application.Current.Resources;var accent=Parse(Accent);var sidebar=Parse(Sidebar);
        void Brush(string key,Color color)=>resources[key]=new SolidColorBrush(color);
        Brush("ThemeAccent",accent);Brush("ThemeAccentForeground",Contrast(accent,Colors.White)>=4.5?Colors.White:Colors.Black);
        Brush("ThemeSidebar",sidebar);Brush("ThemeSidebarSelected",Mix(sidebar,Colors.White,0.10));
        var icon=Mix(accent,Colors.White,0.73);if(Contrast(sidebar,icon)<4.5)icon=Colors.White;
        Brush("ThemeSidebarIcon",icon);Brush("ThemeSidebarMuted",Mix(sidebar,Colors.White,0.78));
        Brush("ThemePage",Accent=="#147D70"&&Sidebar=="#142D3B"?Parse("#F4F7F9"):Mix(Colors.White,accent,0.035));resources["ThemeSidebarColor"]=sidebar;resources["ThemeSidebarTransparentColor"]=Color.FromArgb(0,sidebar.R,sidebar.G,sidebar.B);
    }
}
