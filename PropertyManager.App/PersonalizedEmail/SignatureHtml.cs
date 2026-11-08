using System.Globalization;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml.Linq;
namespace PropertyManager.App.PersonalizedEmail;

internal static class SignatureHtml
{
    internal const string DefaultHtml="<div><p style=\"font-family:Arial;font-size:14px;color:#334155;margin:0 0 4px 0\">Management Office</p><p style=\"font-family:Arial;font-size:18px;color:#147D70;font-weight:bold;margin:0 0 12px 0\">PropertyManagerDemo</p><p style=\"margin:0\"><br /></p><p style=\"font-family:Arial;font-size:12px;color:#64748B;border-top:1px solid #DCE4E8;padding-top:10px;margin:0\">Please do not reply to this email. For assistance, please contact the Management Office.</p></div>";
    internal static string FromPlain(string text)=>new XElement("div",text.Split('\n').Select(s=>new XElement("p",new XAttribute("style","margin:0;font-family:Arial;font-size:14px"),s.TrimEnd('\r').Length==0?new XElement("br"):new XText(s.TrimEnd('\r'))))).ToString(SaveOptions.DisableFormatting);
    internal static XElement Parse(string html)
    {
        if(html.Length>3_000_000)throw new InvalidDataException("Keep a signature below 3 MB, including its embedded pictures.");
        var root=XElement.Parse(html,LoadOptions.PreserveWhitespace);
        var tags=new HashSet<string>{"div","p","span","strong","b","em","i","u","br","img","a","html","head","meta","title","body"};
        foreach(var e in root.DescendantsAndSelf())
        {
            if(!tags.Contains(e.Name.LocalName)||e.Name.NamespaceName.Length>0)throw new InvalidDataException("Unsupported signature HTML element: "+e.Name);
            foreach(var a in e.Attributes())
            {
                if(!new[]{"style","src","alt","width","height","href","charset"}.Contains(a.Name.LocalName))throw new InvalidDataException("Unsupported signature HTML attribute.");
                if(a.Name.LocalName=="style"&&(a.Value.Contains("url",StringComparison.OrdinalIgnoreCase)||a.Value.Contains("expression",StringComparison.OrdinalIgnoreCase)||a.Value.Contains('<')))throw new InvalidDataException("Unsupported signature style.");
            }
            if(e.Name.LocalName=="img")DecodeImage((string?)e.Attribute("src")??"");
            if(e.Name.LocalName=="a"&&!SafeLink((string?)e.Attribute("href")??""))throw new InvalidDataException("Use an https, http or mailto link.");
        }
        return root;
    }
    internal static bool SafeLink(string text)=>Uri.TryCreate(text,UriKind.Absolute,out var uri)&&new[]{"https","http","mailto"}.Contains(uri.Scheme.ToLowerInvariant());
    internal static string PlainText(string html)
    {
        var root=Parse(html);string Text(XElement e)=>e.Name.LocalName=="br"?"\n":e.Name.LocalName=="img"?"":string.Concat(e.Nodes().Select(n=>n is XText t?t.Value:n is XElement c?Text(c):""));
        var paragraphs=root.DescendantsAndSelf().Where(e=>e.Name.LocalName=="p").ToList();
        return (paragraphs.Count>0?string.Join("\n",paragraphs.Select(p=>Text(p).TrimEnd('\n'))):Text(root)).TrimEnd('\r','\n');
    }
    internal static string Message(string body,string signature)
    {
        if(signature.Length>0)Parse(signature);
        return "<!DOCTYPE html><html><head><meta charset=\"utf-8\" /></head><body><div style=\"font-family:Arial;font-size:14px;color:#172B3A;margin:0 0 20px 0\">"+WebUtility.HtmlEncode(body).Replace("\n","<br />")+"</div>"+signature+"</body></html>";
    }
    internal static string Render(string html,Dictionary<string,string> fields)=>CampaignService.Render(html,fields.ToDictionary(p=>p.Key,p=>WebUtility.HtmlEncode(p.Value),StringComparer.OrdinalIgnoreCase));
    internal static string TestBanner(string html,string text)=>html.Replace("<body>","<body><p style=\"font-family:Arial;font-size:14px;color:#795619\">"+WebUtility.HtmlEncode(text).Replace("\n","<br />")+"</p>");
    static Dictionary<string,string> Styles(XElement e)=>((string?)e.Attribute("style")??"").Split(';',StringSplitOptions.RemoveEmptyEntries).Select(v=>v.Split(':',2)).Where(v=>v.Length==2).GroupBy(v=>v[0].Trim()).ToDictionary(g=>g.Key,g=>g.Last()[1].Trim(),StringComparer.OrdinalIgnoreCase);
    static double Num(string s,double fallback)=>double.TryParse(Regex.Match(s,@"[0-9.]+").Value,NumberStyles.Float,CultureInfo.InvariantCulture,out double d)?d:fallback;
    static void Style(TextElement target,XElement e)
    {
        var css=Styles(e);if(css.TryGetValue("font-family",out var family))target.FontFamily=new FontFamily(family.Trim('\''));
        if(css.TryGetValue("font-size",out var size))target.FontSize=Math.Clamp(Num(size,14),8,48);
        if(css.TryGetValue("color",out var color))try{target.Foreground=new SolidColorBrush((Color)ColorConverter.ConvertFromString(color));}catch{}
        if(css.GetValueOrDefault("font-weight")=="bold"||e.Name.LocalName is "b" or "strong")target.FontWeight=FontWeights.Bold;
        if(css.GetValueOrDefault("font-style")=="italic"||e.Name.LocalName is "i" or "em")target.FontStyle=FontStyles.Italic;
        if(target is Span span&&(css.GetValueOrDefault("text-decoration")=="underline"||e.Name.LocalName=="u"))span.TextDecorations=TextDecorations.Underline;
    }
    internal static FlowDocument Document(string html)
    {
        var root=Parse(html);var doc=new FlowDocument{FontFamily=new FontFamily("Arial"),FontSize=14,PagePadding=new Thickness(12)};
        Inline Inline(XNode n)
        {
            if(n is XText t)return new Run(t.Value);
            if(n is not XElement e)return new Run("");
            if(e.Name.LocalName=="br")return new LineBreak();
            if(e.Name.LocalName=="img")
            {
                string source=(string)e.Attribute("src")!;var bitmap=DecodeImage(source);
                return new InlineUIContainer(new Image{Source=bitmap,Width=Math.Clamp(Num((string?)e.Attribute("width")??"160",160),16,800),Stretch=Stretch.Uniform,Tag=source,ToolTip=(string?)e.Attribute("alt")??"Signature picture"});
            }
            Span span=e.Name.LocalName=="a"?new Hyperlink{NavigateUri=new Uri((string)e.Attribute("href")!)}:new Span();Style(span,e);
            foreach(var child in e.Nodes())span.Inlines.Add(Inline(child));return span;
        }
        void Blocks(XElement e)
        {
            if(e.Name.LocalName=="head")return;
            if(e.Name.LocalName is "html" or "body"||e.Name.LocalName=="div"&&e.Elements().Any(c=>c.Name.LocalName is "p" or "div"))
            {foreach(var child in e.Elements())Blocks(child);return;}
            var p=new Paragraph{Margin=new Thickness(0,0,0,4)};Style(p,e);var css=Styles(e);
            if(css.TryGetValue("text-align",out var align)&&Enum.TryParse<TextAlignment>(align,true,out var alignment))p.TextAlignment=alignment;
            if(css.ContainsKey("border-top")){p.BorderThickness=new Thickness(0,1,0,0);p.BorderBrush=Brushes.LightGray;p.Padding=new Thickness(0,10,0,0);}
            foreach(var child in e.Nodes())p.Inlines.Add(Inline(child));doc.Blocks.Add(p);
        }
        Blocks(root);if(doc.Blocks.Count==0)doc.Blocks.Add(new Paragraph());return doc;
    }
    internal static string Html(FlowDocument doc)
    {
        string Css(TextElement e)=>$"font-family:{e.FontFamily.Source};font-size:{e.FontSize.ToString("0.##",CultureInfo.InvariantCulture)}px;color:{(e.Foreground as SolidColorBrush)?.Color.ToString()??"#FF172B3A"};font-weight:{(e.FontWeight.ToOpenTypeWeight()>=600?"bold":"normal")};font-style:{(e.FontStyle==FontStyles.Italic?"italic":"normal")}".Replace("#FF","#");
        object Inline(Inline inline)
        {
            if(inline is LineBreak)return new XElement("br");
            if(inline is InlineUIContainer{Child:Image image})
            {
                string source=image.Tag as string??throw new InvalidDataException("Insert pictures using Add picture so they are embedded in your signature.");DecodeImage(source);
                return new XElement("img",new XAttribute("src",source),new XAttribute("width",(int)(double.IsFinite(image.Width)?image.Width:160)),new XAttribute("alt",image.ToolTip?.ToString()??"Signature picture"));
            }
            if(inline is Run run)return new XElement("span",new XAttribute("style",Css(run)+(run.TextDecorations.Contains(TextDecorations.Underline[0])?";text-decoration:underline":"")),run.Text);
            if(inline is Span span){var node=new XElement(span is Hyperlink?"a":"span",new XAttribute("style",Css(span)+(span.TextDecorations.Contains(TextDecorations.Underline[0])?";text-decoration:underline":"")));if(span is Hyperlink link)node.SetAttributeValue("href",link.NavigateUri?.AbsoluteUri??"");foreach(var c in span.Inlines)node.Add(Inline(c));return node;}
            throw new InvalidDataException("Unsupported signature content. Use text and inserted pictures.");
        }
        var root=new XElement("div");
        foreach(var block in doc.Blocks)
        {
            if(block is not Paragraph p)throw new InvalidDataException("Use paragraphs for signatures; pasted tables/lists are not supported.");
            string style=Css(p)+";margin:0 0 4px 0;text-align:"+p.TextAlignment.ToString().ToLowerInvariant()+(p.BorderThickness.Top>0?";border-top:1px solid #DCE4E8;padding-top:10px":"");
            var paragraph=new XElement("p",new XAttribute("style",style));foreach(var inline in p.Inlines)paragraph.Add(Inline(inline));if(!paragraph.Nodes().Any())paragraph.Add(new XElement("br"));root.Add(paragraph);
        }
        string html=root.ToString(SaveOptions.DisableFormatting);Parse(html);return html;
    }
    internal static BitmapImage DecodeImage(string source)
    {
        const string prefix="data:image/png;base64,";if(!source.StartsWith(prefix,StringComparison.Ordinal))throw new InvalidDataException("Signature pictures must be embedded PNG images.");
        byte[] bytes=Convert.FromBase64String(source[prefix.Length..]);if(bytes.Length>1_000_000)throw new InvalidDataException("Keep each signature picture below 1 MB.");
        var image=new BitmapImage();using(var stream=new MemoryStream(bytes)){image.BeginInit();image.CacheOption=BitmapCacheOption.OnLoad;image.StreamSource=stream;image.EndInit();}if((long)image.PixelWidth*image.PixelHeight>16_000_000)throw new InvalidDataException("Signature picture dimensions are too large.");image.Freeze();return image;
    }
    internal static string ImportPicture(string path)
    {
        if(new FileInfo(path).Length>5_000_000)throw new InvalidDataException("Choose a PNG or JPEG picture below 5 MB.");
        var image=new BitmapImage();using(var stream=File.OpenRead(path)){image.BeginInit();image.CacheOption=BitmapCacheOption.OnLoad;image.DecodePixelWidth=800;image.StreamSource=stream;image.EndInit();}image.Freeze();
        BitmapSource output=image;if(image.PixelHeight>400)output=new TransformedBitmap(image,new ScaleTransform(400.0/image.PixelHeight,400.0/image.PixelHeight));
        var png=new PngBitmapEncoder();png.Frames.Add(BitmapFrame.Create(output));using var bytes=new MemoryStream();png.Save(bytes);if(bytes.Length>1_000_000)throw new InvalidDataException("Compress the picture to below 1 MB before inserting it.");return "data:image/png;base64,"+Convert.ToBase64String(bytes.ToArray());
    }
}
