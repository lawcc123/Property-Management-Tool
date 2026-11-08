using System.Globalization;
using System.Windows;
using System.Windows.Controls;

namespace PropertyManager.App.Envelopes;

public partial class EnvelopeView
{
    private int firstLabelSlot=1;
    private List<int>? firstPageSlots;
    private int labelPage;
    private bool Labels=>OutputModeCombo.SelectedIndex==1;
    private LabelLayout SelectedLabel=>LabelTemplateCombo.SelectedItem as LabelLayout??LabelLayout.All[0];
    private void ConfigureLabelControls(EnvelopeOptions options)
    {
        OutputModeCombo.SelectedIndex=options.IsLabel?1:0;
        IncludeReturnLabels.IsChecked=options.ReturnLabelCount>0;
        ReturnLabelCountText.Text=Math.Max(1,options.ReturnLabelCount).ToString(CultureInfo.InvariantCulture);
        var layouts=LabelLayout.All.ToList();
        if(options.LabelSheet is {} historical&&!layouts.Contains(historical))layouts.Add(historical);
        LabelTemplateCombo.ItemsSource=layouts;
        LabelTemplateCombo.SelectedItem=options.LabelSheet??layouts[0];
        SetLabelPositions(options.FirstLabelSlot,options.FirstPageSlots);
        UpdateOutputControls();labelPage=0;
    }
    private void UpdateOutputControls()
    {
        CreateWordButton.Visibility=Labels?Visibility.Collapsed:Visibility.Visible;
        PrintButton.Visibility=PrinterPropertiesButton.Visibility=PrinterSelectionPanel.Visibility=TraySelectionPanel.Visibility=DriverActionsPanel.Visibility=DefaultBypassCheck.Visibility=DriverInstructions.Visibility=PrinterStatus.Visibility=Labels?Visibility.Visible:Visibility.Collapsed;
        MailingHelpText.Text=Labels?"Preview your labels, choose a paper tray, and keep a history of every request.":"Create an envelope document, then select your printer and Bypass Tray in Word.";
        CancelPrintButton.Content=Labels?"Cancel printing":"Cancel generation";
        LabelSettingsPanel.Visibility=PreviewLabelPositionsButton.Visibility=Labels?Visibility.Visible:Visibility.Collapsed;
        EnvelopeSizePanel.Visibility=Labels?Visibility.Collapsed:Visibility.Visible;
        ReturnSettingsPanel.Visibility=!Labels||IncludeReturnLabels.IsChecked==true?Visibility.Visible:Visibility.Collapsed;
        ReturnLabelQuantityPanel.IsEnabled=IncludeReturnLabels.IsChecked==true;
        ReturnFontText.IsEnabled=!Labels;
        WidthText.IsEnabled=HeightText.IsEnabled=!Labels&&(SizeCombo.SelectedItem as EnvelopePreset)?.Key=="Custom";
        PrintButton.Content="Print selected labels";
        LabelDimensionsText.Text=$"Each label: {SelectedLabel.LabelWidth*25.4:0.##} x {SelectedLabel.LabelHeight*25.4:0.##} mm. Sheet: US Letter, 215.9 x 279.4 mm.";
    }
    private void ReturnLabels_Changed(object s,RoutedEventArgs e)
    {
        if(loading)return;
        if(IncludeReturnLabels.IsChecked==true&&(ReturnCombo.SelectedItem as ReturnChoice)?.Address is null)
        {
            bool previous=loading;loading=true;
            ReturnCombo.SelectedItem=ReturnCombo.Items.OfType<ReturnChoice>().FirstOrDefault(c=>c.Address is not null)??ReturnCombo.SelectedItem;
            loading=previous;
        }
        labelPage=0;imageable=null;UpdateOutputControls();UpdatePreview();
    }
    private int ReadReturnLabelCount()
    {
        if(!Labels||IncludeReturnLabels.IsChecked!=true)return 0;
        if(!int.TryParse(ReturnLabelCountText.Text,out int count)||count<1||count>10000)
            throw new InvalidOperationException("Enter a whole number from 1 to 10,000 for return-address labels.");
        return count;
    }
    private void SetLabelPositions(int first=1,List<int>? specific=null)
    {
        firstLabelSlot=first;firstPageSlots=specific?.ToList();
        LabelPositionSummary.Text=specific is null?$"First sheet: start at position {first}. Later sheets use all positions."
            :"First sheet: use positions "+string.Join(", ",specific.Order())+". Later sheets use all positions.";
    }
    private void LabelPositions_Click(object s,RoutedEventArgs e)
    {
        if(busy||!Labels)return;
        var dialog=new LabelPositionsWindow(SelectedLabel,firstLabelSlot,firstPageSlots){Owner=OwnerWindow};
        if(dialog.ShowDialog()!=true)return;
        SetLabelPositions(dialog.FirstSlot,dialog.SpecificSlots);labelPage=0;imageable=null;UpdatePreview();
    }
    private void OutputMode_Changed(object s,SelectionChangedEventArgs e)
    {
        if(loading)return;
        loading=true;labelPage=0;
        if(Labels){WidthText.Text="215.9";HeightText.Text="279.4";FontText.Text=SelectedLabel.Columns==3?"10":"14";}
        else
        {
            var preset=SizeCombo.SelectedItem as EnvelopePreset??EnvelopePreset.All[0];
            WidthText.Text=preset.WidthMm.ToString(CultureInfo.InvariantCulture);HeightText.Text=preset.HeightMm.ToString(CultureInfo.InvariantCulture);FontText.Text="14";
        }
        UpdateOutputControls();loading=false;
        if(printer is WindowsEnvelopePrinter windows)windows.ClearDriverPreferences();
        LoadProfile();imageable=null;UpdatePreview();
    }
    private void LabelTemplate_Changed(object s,SelectionChangedEventArgs e)
    {
        if(loading)return;
        loading=true;SetLabelPositions();FontText.Text=SelectedLabel.Columns==3?"10":"14";labelPage=0;UpdateOutputControls();loading=false;
        LoadProfile();imageable=null;UpdatePreview();
    }
    private EnvelopeDrawing CurrentDrawing(bool guides)
    {
        if(!Labels)
        {
            var address=CurrentAddress??PreviewPlaceholder;
            return EnvelopeRenderer.Render(address,(ReturnCombo.SelectedItem as ReturnChoice)?.Address,ReadOptions(),imageable,guides);
        }
        var snapshot=Snapshot(preview:true);labelPage=Math.Clamp(labelPage,0,snapshot.Pages-1);
        return MailingRenderer.RenderPage(snapshot,labelPage,imageable,guides);
    }
    private string CurrentCaption()
    {
        var options=ReadOptions();
        if(CurrentAddress is null&&(!Labels||options.ReturnLabelCount==0))return "Layout preview - select contractors to preview their addresses";
        if(!Labels)return $"{CurrentAddress?.Company} \u00B7 {options.WidthMm:0.#} \u00D7 {options.HeightMm:0.#} mm";
        var snapshot=Snapshot(preview:true);
        return $"{SelectedLabel.Label} - sheet {labelPage+1} of {snapshot.Pages} - {snapshot.Items} label(s) - US Letter";
    }
}
