using System.Globalization;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace PropertyManager.App.Envelopes;

internal sealed record EnvelopeDrawing(DrawingGroup Drawing,Rect Page,Rect ContentBounds,List<string> Warnings);
internal static class EnvelopeRenderer
{
    internal const double DipsPerMm=96/25.4;
    internal static EnvelopeDrawing Render(MailingAddress recipient,MailingAddress? sender,EnvelopeOptions options,Rect? imageable=null,bool guides=false)
    {
        options.Validate();
        if(options.IncludeReturnLogo&&sender?.LogoPng is null)throw new InvalidOperationException("Select a return address with a saved logo.");
        var page=new Rect(0,0,options.WidthMm*DipsPerMm,options.HeightMm*DipsPerMm);
        var group=new DrawingGroup();var warnings=new List<string>();Rect bounds=Rect.Empty;Rect returnBounds=Rect.Empty;
        using(var dc=group.Open())
        {
            dc.DrawRectangle(Brushes.White,null,page);
            if(options.Rotate180)dc.PushTransform(new RotateTransform(180,page.Width/2,page.Height/2));
            double dx=options.OffsetXmm*DipsPerMm,dy=options.OffsetYmm*DipsPerMm;
            if(sender is not null)
            {
                if(options.IncludeReturnLogo)
                    returnBounds=ReturnAddressLogo.Draw(dc,sender,options.ReturnFontPt,new Rect(9*DipsPerMm+dx,8*DipsPerMm+dy,Math.Max(1,page.Width-18*DipsPerMm),Math.Max(1,page.Height-16*DipsPerMm)),false);
                else
                {
                    var text=Text(sender.Text,options.ReturnFontPt,Math.Max(1,page.Width-18*DipsPerMm));
                    returnBounds=new Rect(9*DipsPerMm+dx,8*DipsPerMm+dy,text.WidthIncludingTrailingWhitespace,text.Height);
                    dc.DrawText(text,returnBounds.TopLeft);
                }
                bounds.Union(returnBounds);
            }
            var address=Text(recipient.Text,options.FontPt,Math.Max(1,page.Width-28*DipsPerMm));
            address.TextAlignment=TextAlignment.Center;
            double top=page.Height*0.54-address.Height/2;
            var recipientRect=new Rect(14*DipsPerMm+dx,top+dy,page.Width-28*DipsPerMm,address.Height);
            dc.DrawText(address,recipientRect.TopLeft);bounds.Union(address.BuildGeometry(recipientRect.TopLeft).Bounds);
            if(sender is not null)
            {
                if(recipientRect.Top<returnBounds.Bottom+4*DipsPerMm)warnings.Add("The return and recipient addresses overlap. Reduce text size or use a larger envelope.");
            }
            if(options.Rotate180){dc.Pop();if(!bounds.IsEmpty)bounds=new Rect(page.Width-bounds.Right,page.Height-bounds.Bottom,bounds.Width,bounds.Height);}
            if(!bounds.IsEmpty&&!page.Contains(bounds))warnings.Add("Address text extends beyond the envelope. Adjust the layout before printing.");
            if(imageable is Rect printable&&!bounds.IsEmpty&&!printable.Contains(bounds))warnings.Add("Address text is outside this printer's printable area. Adjust alignment or text size.");
            if(guides)
            {
                var guide=imageable??new Rect(5*DipsPerMm,5*DipsPerMm,Math.Max(1,page.Width-10*DipsPerMm),Math.Max(1,page.Height-10*DipsPerMm));
                dc.DrawRectangle(null,new Pen(Brushes.SlateGray,0.7){DashStyle=DashStyles.Dash},guide);
                dc.DrawRectangle(null,new Pen(Brushes.LightGray,1),page);
            }
        }
        group.Freeze();return new(group,page,bounds,warnings);
    }
    internal static FormattedText Text(string value,double fontPt,double width)
    {
        var text=new FormattedText(value,CultureInfo.GetCultureInfo("en-CA"),FlowDirection.LeftToRight,
            new Typeface(new FontFamily("Arial"),FontStyles.Normal,FontWeights.Bold,FontStretches.Normal),fontPt*96/72,Brushes.Black,1);
        text.MaxTextWidth=width;text.Trimming=TextTrimming.None;return text;
    }
    internal static byte[] Png(EnvelopeDrawing drawing,int dpi=300)
    {
        var visual=new DrawingVisual();using(var dc=visual.RenderOpen())dc.DrawDrawing(drawing.Drawing);
        var bitmap=new RenderTargetBitmap((int)Math.Ceiling(drawing.Page.Width*dpi/96),(int)Math.Ceiling(drawing.Page.Height*dpi/96),dpi,dpi,PixelFormats.Pbgra32);
        bitmap.Render(visual);var png=new PngBitmapEncoder();png.Frames.Add(BitmapFrame.Create(bitmap));
        using var buffer=new System.IO.MemoryStream();png.Save(buffer);return buffer.ToArray();
    }
}
