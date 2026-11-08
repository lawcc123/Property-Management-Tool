using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
namespace PropertyManager.App.PersonalizedEmail;

// A text-only projection of Word fields, not a Word layout/field-expression engine.
// Native MERGEFIELD cached results must never masquerade as current recipient values.
internal static class WordPreviewText
{
    private static readonly XNamespace W="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly Regex Merge=new(@"^\s*MERGEFIELD\s+(?:""(?<name>[^""]+)""|(?<name>[^\s\\]+))",RegexOptions.IgnoreCase);
    private sealed class Field
    {
        internal readonly StringBuilder Code=new();
        internal bool Result;
        internal string? Name;
    }
    internal static string Read(XDocument document)
    {
        var text=new StringBuilder();var fields=new Stack<Field>();
        bool Visible()=>fields.All(f=>f.Result&&f.Name is null);
        string? Name(string code){var match=Merge.Match(code);return match.Success?match.Groups["name"].Value:null;}
        void Token(string name){if(Visible())text.Append("{{").Append(name).Append("}}");}
        void Visit(XElement element)
        {
            if(element.Name==W+"fldSimple")
            {
                var name=Name((string?)element.Attribute(W+"instr")??"");
                if(name is not null){Token(name);return;}
            }
            if(element.Name==W+"fldChar")
            {
                string? type=(string?)element.Attribute(W+"fldCharType");
                if(type=="begin")fields.Push(new());
                else if(type=="separate"&&fields.TryPeek(out var field))
                {
                    field.Name=Name(field.Code.ToString());
                    // Emit outside this field so only enclosing fields control visibility.
                    fields.Pop();if(field.Name is not null)Token(field.Name);fields.Push(field);field.Result=true;
                }
                else if(type=="end"&&fields.TryPop(out var ended)&&!ended.Result)
                {
                    var name=Name(ended.Code.ToString());if(name is not null)Token(name);
                }
                return;
            }
            if(element.Name==W+"instrText")
            {
                if(fields.TryPeek(out var field)&&!field.Result)field.Code.Append(element.Value);
                return;
            }
            if(element.Name==W+"t") {if(Visible())text.Append(element.Value);return;}
            if(element.Name==W+"tab") {if(Visible())text.Append('\t');return;}
            if(element.Name==W+"br"||element.Name==W+"cr") {if(Visible())text.AppendLine();return;}
            foreach(var child in element.Elements())Visit(child);
            if(element.Name==W+"p")text.AppendLine();
        }
        if(document.Root is not null)Visit(document.Root);
        if(fields.Count>0)throw new InvalidDataException("This Word template contains an incomplete field. Open and repair the field in Word before previewing.");
        return text.ToString();
    }
}
