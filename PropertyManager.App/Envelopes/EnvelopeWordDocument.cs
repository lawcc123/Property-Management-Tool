using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;

namespace PropertyManager.App.Envelopes;

// A portable .docx, built without Word automation or printer access.
internal static class EnvelopeWordDocument
{
    static readonly XNamespace W="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly XNamespace R="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    static int Twips(double mm)=>(int)Math.Round(mm*1440/25.4);
    static XElement E(string name,params object[] content)=>new(W+name,content);
    static XAttribute A(string name,object value)=>new(W+name,value);
    static XElement Paragraph(string text,double size,double line,bool center=false,bool pageBreak=false,double indent=0)
        =>E("p",E("pPr",E("spacing",A("before",0),A("after",0),A("line",(int)Math.Round(line*20)),A("lineRule","exact")),E("jc",A("val",center?"center":"left")),E("ind",A("left",Twips(indent))),pageBreak?E("pageBreakBefore"):null!),
            E("r",E("rPr",E("rFonts",A("ascii","Arial"),A("hAnsi","Arial")),E("b"),E("sz",A("val",(int)Math.Round(size*2)))),E("t",new XAttribute(XNamespace.Xml+"space","preserve"),text)));
    internal static void Write(string path,EnvelopeSnapshot snapshot,Action<int,int>? progress=null,CancellationToken cancel=default)
    {
        snapshot.Validate();if(snapshot.Options.IsLabel)throw new InvalidOperationException("The Word option is for envelopes. Use direct printing for labels.");
        var o=snapshot.Options;var body=E("body");var images=new List<byte[]>();
        bool imageLayout=o.IncludeReturnLogo||o.Rotate180;
        for(int page=0;page<snapshot.Pages;page++)
        {
            cancel.ThrowIfCancellationRequested();var address=snapshot.Recipients[page/o.Copies];
            if(imageLayout)
            {
                // Preserve logo/rotation exactly without requiring Word to render a custom text box.
                images.Add(EnvelopeRenderer.Png(EnvelopeRenderer.Render(address,snapshot.ReturnAddress,o),300));
                body.Add(ImageParagraph(page,o,page>0));
            }
            else
            {
                var returns=snapshot.ReturnAddress?.Text.Split('\n').Where(s=>s.Length>0).ToArray()??[];
                int n=0;
                foreach(string line in returns)body.Add(Paragraph(line.TrimEnd('\r'),o.ReturnFontPt,o.ReturnFontPt+3,pageBreak:page>0&&n++==0,indent:o.OffsetXmm));
                body.Add(Paragraph("",1,Math.Max(1,106-6*returns.Length+o.OffsetYmm*72/25.4),pageBreak:page>0&&returns.Length==0));
                foreach(string line in address.Text.Split('\n'))body.Add(Paragraph(line.TrimEnd('\r'),o.FontPt,o.FontPt+4,true,indent:o.OffsetXmm*2));
            }
            progress?.Invoke(page+1,snapshot.Pages);
        }
        // Matches the original Python project's custom envelope page and margins.
        // Do not embed COM10, Letter, a printer bin or a driver-specific DEVMODE.
        body.Add(E("sectPr",E("pgSz",A("w",Twips(o.WidthMm)),A("h",Twips(o.HeightMm))),
            E("pgMar",A("top",Twips(8)),A("bottom",Twips(5)),A("left",Twips(9)),A("right",Twips(5)),A("header",0),A("footer",0),A("gutter",0))));
        using var file=new FileStream(path,FileMode.CreateNew,FileAccess.Write,FileShare.None);
        using var zip=new ZipArchive(file,ZipArchiveMode.Create);
        void Xml(string name,XElement element){using var stream=zip.CreateEntry(name).Open();new XDocument(element).Save(stream);}
        XNamespace ct="http://schemas.openxmlformats.org/package/2006/content-types", rel="http://schemas.openxmlformats.org/package/2006/relationships";
        Xml("[Content_Types].xml",new XElement(ct+"Types",new XElement(ct+"Default",new XAttribute("Extension","rels"),new XAttribute("ContentType","application/vnd.openxmlformats-package.relationships+xml")),new XElement(ct+"Default",new XAttribute("Extension","png"),new XAttribute("ContentType","image/png")),new XElement(ct+"Override",new XAttribute("PartName","/word/document.xml"),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"))));
        Xml("_rels/.rels",new XElement(rel+"Relationships",new XElement(rel+"Relationship",new XAttribute("Id","rId1"),new XAttribute("Type",R.NamespaceName+"/officeDocument"),new XAttribute("Target","word/document.xml"))));
        Xml("word/document.xml",new XElement(W+"document",new XAttribute(XNamespace.Xmlns+"w",W),new XAttribute(XNamespace.Xmlns+"r",R),body));
        Xml("word/_rels/document.xml.rels",new XElement(rel+"Relationships",images.Select((_,i)=>new XElement(rel+"Relationship",new XAttribute("Id","image"+i),new XAttribute("Type",R.NamespaceName+"/image"),new XAttribute("Target",$"media/page{i}.png")))));
        for(int i=0;i<images.Count;i++){using var stream=zip.CreateEntry($"word/media/page{i}.png").Open();stream.Write(images[i]);}
    }
    static XElement ImageParagraph(int index,EnvelopeOptions o,bool pageBreak)
    {
        XNamespace wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", a="http://schemas.openxmlformats.org/drawingml/2006/main", pic="http://schemas.openxmlformats.org/drawingml/2006/picture";
        long cx=(long)Math.Round(o.WidthMm*36000),cy=(long)Math.Round(o.HeightMm*36000);
        return E("p",E("pPr",E("spacing",A("before",0),A("after",0),A("line",20),A("lineRule","exact")),pageBreak?E("pageBreakBefore"):null!),E("r",E("drawing",
            new XElement(wp+"anchor",new XAttribute("distT",0),new XAttribute("distB",0),new XAttribute("distL",0),new XAttribute("distR",0),new XAttribute("simplePos",0),new XAttribute("relativeHeight",0),new XAttribute("behindDoc",0),new XAttribute("locked",0),new XAttribute("layoutInCell",1),new XAttribute("allowOverlap",1),
            new XElement(wp+"simplePos",new XAttribute("x",0),new XAttribute("y",0)),new XElement(wp+"positionH",new XAttribute("relativeFrom","page"),new XElement(wp+"posOffset",0)),new XElement(wp+"positionV",new XAttribute("relativeFrom","page"),new XElement(wp+"posOffset",0)),new XElement(wp+"extent",new XAttribute("cx",cx),new XAttribute("cy",cy)),new XElement(wp+"wrapNone"),new XElement(wp+"docPr",new XAttribute("id",index+1),new XAttribute("name","Envelope "+(index+1))),
            new XElement(a+"graphic",new XElement(a+"graphicData",new XAttribute("uri",pic.NamespaceName),new XElement(pic+"pic",new XElement(pic+"nvPicPr",new XElement(pic+"cNvPr",new XAttribute("id",index+1),new XAttribute("name","Envelope")),new XElement(pic+"cNvPicPr")),new XElement(pic+"blipFill",new XElement(a+"blip",new XAttribute(R+"embed","image"+index)),new XElement(a+"stretch",new XElement(a+"fillRect"))),new XElement(pic+"spPr",new XElement(a+"xfrm",new XElement(a+"off",new XAttribute("x",0),new XAttribute("y",0)),new XElement(a+"ext",new XAttribute("cx",cx),new XAttribute("cy",cy))),new XElement(a+"prstGeom",new XAttribute("prst","rect"),new XElement(a+"avLst"))))))))));
    }
}
