using System.IO;
using System.Text;
using Microsoft.VisualBasic.FileIO;

namespace PropertyManager.App.Envelopes;

internal static class AddressCsv
{
    internal static List<MailingAddress> Read(string path)
    {
        using var csv=new TextFieldParser(path,Encoding.UTF8,true){TextFieldType=FieldType.Delimited,HasFieldsEnclosedInQuotes=true,TrimWhiteSpace=false};csv.SetDelimiters(",");
        var headers=(csv.ReadFields()??[]).Select(h=>h.Trim().ToLowerInvariant()).ToList();
        if(headers.Count!=headers.Distinct().Count()||headers.Any(h=>h.Length==0)||new[]{"company","address","city"}.Any(h=>!headers.Contains(h)))
            throw new InvalidDataException("CSV needs unique column headings including company, address, and city. Optional fields: id, nickname, province, postal.");
        var addresses=new List<MailingAddress>();
        while(!csv.EndOfData)
        {
            var values=csv.ReadFields()??[];if(values.All(string.IsNullOrWhiteSpace))continue;
            if(values.Length>headers.Count)throw new InvalidDataException("A CSV row has more values than its headings.");
            string Field(string key){int i=headers.IndexOf(key);return i>=0&&i<values.Length?values[i].Trim():"";}
            var address=new MailingAddress{Id=string.IsNullOrWhiteSpace(Field("id"))?Guid.NewGuid().ToString("N"):Field("id"),Company=Field("company"),Street=Field("address"),City=Field("city"),Province=Field("province"),Postal=Field("postal"),Nickname=Field("nickname")};
            address.Validate();addresses.Add(address);
        }
        return addresses;
    }
    internal static string Write(IEnumerable<MailingAddress> addresses)
    {
        static string Quote(string text)=>"\""+((text.Length>0&&"=+-@\t\r".Contains(text[0]))?"'":"")+text.Replace("\"","\"\"")+"\"";
        var output=new StringBuilder("id,company,address,city,province,postal,nickname\r\n");
        foreach(var address in addresses)output.AppendLine(string.Join(",",new[]{address.Id,address.Company,address.Street,address.City,address.Province,address.Postal,address.Nickname}.Select(Quote)));
        return output.ToString();
    }
    internal static int Merge(List<MailingAddress> existing,IEnumerable<MailingAddress> incoming)
    {
        // Work on a copy so an invalid ID later in the input cannot partly import a file.
        var merged=existing.ToList();int added=0;
        foreach(var address in incoming)
        {
            if(merged.Any(a=>a.Company.Equals(address.Company,StringComparison.OrdinalIgnoreCase)&&a.Street.Equals(address.Street,StringComparison.OrdinalIgnoreCase)))continue;
            if(merged.Any(a=>a.Id==address.Id))throw new InvalidDataException("An imported address ID is already used by a different address. No addresses were imported.");
            merged.Add(address);added++;
        }
        existing.Clear();existing.AddRange(merged);return added;
    }
}
