using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace PropertyManager.App.Envelopes;

internal static class ReturnAddressLogo
{
    private const int MaxPngBytes=2*1024*1024;
    private static readonly ConditionalWeakTable<string,BitmapSource> Cache=new();
    internal static string Import(string path)
    {
        if(new FileInfo(path).Length>10*1024*1024)throw new InvalidDataException("Choose an image smaller than 10 MB.");
        using var file=File.OpenRead(path);
        var decoder=BitmapDecoder.Create(file,BitmapCreateOptions.PreservePixelFormat,BitmapCacheOption.None);
        if(decoder is not (PngBitmapDecoder or JpegBitmapDecoder or BmpBitmapDecoder))
            throw new InvalidDataException("Choose a PNG, JPEG or BMP image.");
        var frame=decoder.Frames[0];
        if(frame.PixelWidth>8192||frame.PixelHeight>8192||(long)frame.PixelWidth*frame.PixelHeight>20000000)
            throw new InvalidDataException("Choose a smaller image (up to 20 megapixels).");
        double scale=Math.Min(1,1600.0/Math.Max(frame.PixelWidth,frame.PixelHeight));
        BitmapSource image=scale<1?new TransformedBitmap(frame,new ScaleTransform(scale,scale)):frame;
        var encoder=new PngBitmapEncoder();encoder.Frames.Add(BitmapFrame.Create(image));
        using var output=new MemoryStream();encoder.Save(output);
        if(output.Length>MaxPngBytes)throw new InvalidDataException("The logo is too large. Choose a simpler or smaller image.");
        string data=Convert.ToBase64String(output.ToArray());Decode(data);return data;
    }
    internal static BitmapSource Decode(string data)=>Cache.GetValue(data,value=>
    {
        try
        {
            if(value.Length>MaxPngBytes*4/3+4)throw new InvalidDataException("The saved logo is too large.");
            using var stream=new MemoryStream(Convert.FromBase64String(value));
            var decoder=new PngBitmapDecoder(stream,BitmapCreateOptions.PreservePixelFormat,BitmapCacheOption.OnLoad);
            var image=decoder.Frames[0];
            if(image.PixelWidth>1600||image.PixelHeight>1600)throw new InvalidDataException("The saved logo dimensions are too large.");
            // A frozen BitmapFrame can still retain its thread-owned decoder.
            // Copy decoded pixels into an independent source before caching it.
            var pixelsImage=new FormatConvertedBitmap(image,PixelFormats.Pbgra32,null,0);
            int stride=checked(image.PixelWidth*4);var pixels=new byte[checked(stride*image.PixelHeight)];pixelsImage.CopyPixels(pixels,stride,0);
            var detached=BitmapSource.Create(image.PixelWidth,image.PixelHeight,96,96,PixelFormats.Pbgra32,null,pixels,stride);
            detached.Freeze();return detached;
        }
        catch(Exception ex) when(ex is FormatException or NotSupportedException or ArgumentException or IOException)
        {throw new InvalidOperationException("The saved return-address logo is invalid. Choose the image again.",ex);}
    });
    internal static Rect Draw(DrawingContext dc,MailingAddress sender,double fontPt,Rect area,bool centered)
    {
        var image=Decode(sender.LogoPng!);double mm=EnvelopeRenderer.DipsPerMm;
        double width=Math.Min(sender.LogoWidthMm*mm,area.Width*0.32);
        double height=width*image.PixelHeight/image.PixelWidth;
        double maxHeight=Math.Min(16*mm,area.Height);
        if(height>maxHeight){width*=maxHeight/height;height=maxHeight;}
        double gap=3*mm;
        var text=EnvelopeRenderer.Text(sender.Text,fontPt,Math.Max(1,area.Width-width-gap));
        double totalWidth=width+gap+text.WidthIncludingTrailingWhitespace,totalHeight=Math.Max(height,text.Height);
        double x=area.X+(centered?(area.Width-totalWidth)/2:0),y=area.Y+(centered?(area.Height-totalHeight)/2:0);
        var logo=new Rect(x,y+(totalHeight-height)/2,width,height);
        var origin=new Point(x+width+gap,y+(totalHeight-text.Height)/2);
        dc.DrawImage(image,logo);dc.DrawText(text,origin);
        var bounds=text.BuildGeometry(origin).Bounds;bounds.Union(logo);return bounds;
    }
}
