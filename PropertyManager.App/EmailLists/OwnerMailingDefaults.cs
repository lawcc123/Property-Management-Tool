using System.Text.RegularExpressions;
namespace PropertyManager.App.EmailLists;
internal sealed record OwnerMailingDefault(string Corporation,string Street,string City,string Postal,bool Townhouse)
{
    public string Label=>Corporation;
    internal static readonly OwnerMailingDefault[] All=[new("DEMO 1001","100 Example Ave","Example City","A1A 1A1",false),new("DEMO 2002","200 Example Ave","Example City","A1A 1A2",true)];
    internal static AlternateMailingAddress Apply(AlternateMailingAddress address,string corporation,string unit,string ownerName)
    {
        // A partly entered alternate address is retained for review, never mixed
        // with property-address fragments. Missing unit records are not blank alternates.
        if(address.Text.Length>0||address.Status.StartsWith("Unit not found",StringComparison.Ordinal))return address;
        var property=All.SingleOrDefault(p=>p.Corporation==corporation);
        if(property is null)return address with {Status="Default address is not configured for this corporation - review"};
        unit=unit.Trim();if(unit.Length==0)return address with {Status="Missing unit number - review"};
        string label=property.Townhouse?"TH "+Regex.Replace(unit,@"^TH[\s-]*","",RegexOptions.IgnoreCase):unit;
        return address with{Name=address.Name.Length>0?address.Name:ownerName,Line1=label+" - "+property.Street,Line2="",City=property.City,Province="ON",Postal=property.Postal,Country="",Status="Default property address ("+corporation+")"};
    }
}
