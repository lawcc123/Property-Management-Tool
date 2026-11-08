using System.Windows;
using System.Windows.Media;

namespace PropertyManager.App.Envelopes;

// Dimensions are inches on US Letter sheets; geometry is stored in every history
// snapshot so later template changes cannot silently move an old print request.
internal sealed record LabelLayout(string Key,string Label,int Columns,int Rows,double LabelWidth,double LabelHeight,
    double Left,double Top,double HorizontalPitch,double VerticalPitch)
{
    public int Capacity=>checked(Columns*Rows);
    internal static readonly LabelLayout[] All =
    [new("5162","Avery 5162 - 2 columns x 7 rows (14 labels)",2,7,4,4.0/3,0.15625,5.0/6,4.1875,4.0/3),
     new("5159","Avery 5159 - 2 columns x 7 rows (14 labels)",2,7,4,1.5,0.15625,0.25,4.1875,1.5),
     new("5160","Avery 5160 / 8160 - 3 columns x 10 rows (30 labels)",3,10,2.625,1,0.1875,0.5,2.75,1),
     new("5161","Avery 5161 - 2 columns x 10 rows (20 labels)",2,10,4,1,0.15625,0.5,4.1875,1),
     new("5163","Avery 5163 - 2 columns x 5 rows (10 labels)",2,5,4,2,0.15625,0.5,4.1875,2)];
    internal void Validate()
    {
        if(Columns<1||Rows<1||Columns>10||Rows>30||string.IsNullOrWhiteSpace(Key)||
            new[]{LabelWidth,LabelHeight,Left,Top,HorizontalPitch,VerticalPitch}.Any(n=>!double.IsFinite(n)||n<0)||
            LabelWidth<0.5||LabelHeight<0.5||HorizontalPitch<LabelWidth||VerticalPitch<LabelHeight||
            Left+(Columns-1)*HorizontalPitch+LabelWidth>8.5+0.001||Top+(Rows-1)*VerticalPitch+LabelHeight>11+0.001)
            throw new InvalidOperationException("The saved label-sheet geometry is invalid or extends off the sheet.");
    }
    internal Rect Cell(int slot)=>new((Left+(slot-1)%Columns*HorizontalPitch)*96,(Top+(slot-1)/Columns*VerticalPitch)*96,LabelWidth*96,LabelHeight*96);
}
internal sealed record LabelPlacement(int Slot,int RecipientIndex);
internal static class MailingRenderer
{
    internal static List<LabelPlacement> Placements(EnvelopeSnapshot snapshot,int page)
    {
        var o=snapshot.Options;var sheet=o.LabelSheet??throw new InvalidOperationException("Choose a label template.");
        if(page<0||page>=snapshot.Pages)throw new ArgumentOutOfRangeException(nameof(page));
        var first=o.FirstPageSlots?.Order().ToArray()??Enumerable.Range(o.FirstLabelSlot,sheet.Capacity-o.FirstLabelSlot+1).ToArray();
        int offset=page==0?0:first.Length+(page-1)*sheet.Capacity;
        var slots=page==0?first:Enumerable.Range(1,sheet.Capacity).ToArray();
        return slots.Take(Math.Max(0,snapshot.Items-offset)).Select((slot,index)=>new LabelPlacement(slot,offset+index<snapshot.ContractorItems?(offset+index)/o.Copies:-1)).ToList();
    }
    internal static EnvelopeDrawing RenderPage(EnvelopeSnapshot snapshot,int page,Rect? imageable=null,bool guides=false)
    {
        if(!snapshot.Options.IsLabel)return EnvelopeRenderer.Render(snapshot.Recipients[page/snapshot.Options.Copies],snapshot.ReturnAddress,snapshot.Options,imageable,guides);
        snapshot.Validate();var o=snapshot.Options;var sheet=o.LabelSheet!;
        var paper=new Rect(0,0,8.5*96,11*96);var group=new DrawingGroup();var warnings=new List<string>();Rect bounds=Rect.Empty;
        using(var dc=group.Open())
        {
            dc.DrawRectangle(Brushes.White,null,paper);
            if(o.Rotate180)dc.PushTransform(new RotateTransform(180,paper.Width/2,paper.Height/2));
            foreach(var placement in Placements(snapshot,page))
            {
                var cell=sheet.Cell(placement.Slot);var area=cell;area.Inflate(-2.5*EnvelopeRenderer.DipsPerMm,-1.5*EnvelopeRenderer.DipsPerMm);
                var recipient=placement.RecipientIndex<0?snapshot.ReturnAddress!:snapshot.Recipients[placement.RecipientIndex];
                if(placement.RecipientIndex<0&&o.IncludeReturnLogo)
                {
                    area.Offset(o.OffsetXmm*EnvelopeRenderer.DipsPerMm,o.OffsetYmm*EnvelopeRenderer.DipsPerMm);
                    var artwork=ReturnAddressLogo.Draw(dc,recipient,o.FontPt,area,true);
                    if(!cell.Contains(artwork))warnings.Add($"Slot {placement.Slot}: return-address logo or text extends outside the label. Reduce font or logo size.");
                    bounds.Union(artwork);continue;
                }
                var text=EnvelopeRenderer.Text(recipient.Text,o.FontPt,area.Width);text.TextAlignment=TextAlignment.Center;
                var origin=new Point(area.X+o.OffsetXmm*EnvelopeRenderer.DipsPerMm,cell.Y+(cell.Height-text.Height)/2+o.OffsetYmm*EnvelopeRenderer.DipsPerMm);
                var textBounds=text.BuildGeometry(origin).Bounds;
                if(!cell.Contains(textBounds))warnings.Add($"Slot {placement.Slot} ({recipient.Company}): text extends outside the label. Reduce font size or adjust alignment.");
                bounds.Union(textBounds);dc.DrawText(text,origin);
            }
            if(o.Rotate180){dc.Pop();if(!bounds.IsEmpty)bounds=new Rect(paper.Width-bounds.Right,paper.Height-bounds.Bottom,bounds.Width,bounds.Height);}
            if(!bounds.IsEmpty&&!paper.Contains(bounds))warnings.Add("Label text extends off the paper.");
            if(imageable is Rect printable&&!bounds.IsEmpty&&!printable.Contains(bounds))warnings.Add("Label text extends outside the printer's printable area. Check alignment and tray settings.");
            if(guides)
            {
                // Guides remain at physical stock positions even when artwork rotates.
                for(int slot=1;slot<=sheet.Capacity;slot++)
                {
                    var cell=sheet.Cell(slot);dc.DrawRectangle(null,new Pen(Brushes.LightSlateGray,0.5){DashStyle=DashStyles.Dash},cell);
                    var number=EnvelopeRenderer.Text(slot.ToString(),6,cell.Width);number.SetForegroundBrush(Brushes.SlateGray);dc.DrawText(number,new Point(cell.X+2,cell.Y+1));
                }
                if(imageable is Rect guide)dc.DrawRectangle(null,new Pen(Brushes.SlateGray,0.6){DashStyle=DashStyles.Dot},guide);
            }
        }
        group.Freeze();return new(group,paper,bounds,warnings.Distinct().ToList());
    }
}
