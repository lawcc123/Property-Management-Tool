using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using Microsoft.Win32;

namespace PropertyManager.App.Envelopes;

public partial class EnvelopeView : UserControl
{
    private readonly EnvelopeStore store;
    private readonly IEnvelopePrinter printer;
    private AddressBook book=new();
    private List<RecipientRow> rows=[];
    private List<EnvelopeHistoryEntry> history=[];
    private ICollectionView? recipientsView;
    private bool loading=true,historySource,busy,loaded;
    private string? parentHistory;
    private string desiredPrinter="",desiredTrayName="";
    private int? desiredTrayId;
    private int printerRevision;
    private PrinterInfo? printerInfo;
    private Rect? imageable;
    private CancellationTokenSource? cancellation;
    private string? addressBookError;
    internal bool IsBusy=>busy;
    internal sealed record ReturnChoice(string Label,MailingAddress? Address);
    private Window OwnerWindow=>Window.GetWindow(this);
    internal EnvelopeView(EnvelopeStore? store=null,IEnvelopePrinter? printer=null)
    {
        this.store=store??new();this.printer=printer??new WindowsEnvelopePrinter();InitializeComponent();
        try{book=this.store.LoadBook();}
        catch(Exception ex){addressBookError=ex.Message;StatusText.Text="The address book could not be loaded: "+ex.Message;}
        SizeCombo.ItemsSource=EnvelopePreset.All;LoadReturnChoices(null);
        EnvelopeOptions options=new();try{options=this.store.LoadOptions();options.Validate();}catch(Exception ex){StatusText.Text="Saved layout unavailable. Defaults are shown. "+ex.Message;options=new();}
        ApplyOptions(options);LoadRows(book.Contractors,[]);loading=false;UpdatePreview();
        Loaded+=async(_,_)=>{if(loaded)return;loaded=true;await RefreshPrintersAsync(false);};
    }
    internal void Cancel(){cancellation?.Cancel();}
    private void Error(Exception ex)=>MessageBox.Show(OwnerWindow,ex.Message,"Printing contractor's address",MessageBoxButton.OK,MessageBoxImage.Warning);
    private static double Number(TextBox box,string label)
    {
        if(!double.TryParse(box.Text,NumberStyles.Float,CultureInfo.CurrentCulture,out double value)&&!double.TryParse(box.Text,NumberStyles.Float,CultureInfo.InvariantCulture,out value))
            throw new InvalidOperationException("Enter a number for "+label+".");return value;
    }
    private EnvelopeOptions ReadOptions()
    {
        if(!int.TryParse(CopiesText.Text,out int copies))throw new InvalidOperationException("Enter a whole number of copies.");
        int returnLabels=ReadReturnLabelCount();
        var tray=TrayCombo.SelectedItem as TrayChoice;
        bool logo=IncludeReturnLogo.IsChecked==true&&(!Labels||returnLabels>0)&&(ReturnCombo.SelectedItem as ReturnChoice)?.Address?.LogoPng is not null;
        var options=new EnvelopeOptions{DefaultToBypass=DefaultBypassCheck.IsChecked==true,IncludeReturnLogo=logo,ReturnLabelCount=returnLabels,LayoutVersion=logo?4:Labels?(returnLabels>0?3:2):1,LabelSheet=Labels?SelectedLabel:null,FirstLabelSlot=Labels?firstLabelSlot:1,FirstPageSlots=Labels?firstPageSlots?.ToList():null,Preset=Labels?"Label-"+SelectedLabel.Key:(SizeCombo.SelectedItem as EnvelopePreset)?.Key??"10",WidthMm=Number(WidthText,"width"),HeightMm=Number(HeightText,"height"),FontPt=Number(FontText,"recipient font"),ReturnFontPt=Number(ReturnFontText,"return font"),OffsetXmm=Number(OffsetXText,"horizontal alignment"),OffsetYmm=Number(OffsetYText,"vertical alignment"),Copies=copies,Rotate180=RotateCheck.IsChecked==true,Printer=PrinterCombo.SelectedItem as string??"",TrayId=tray?.Id,TrayName=tray?.Name??"",ReturnAddressId=(ReturnCombo.SelectedItem as ReturnChoice)?.Address?.Id??""};
        options.Validate();return options;
    }
    private void ApplyOptions(EnvelopeOptions options)
    {
        bool previous=loading;loading=true;
        ConfigureLabelControls(options);
        IncludeReturnLogo.IsChecked=options.IncludeReturnLogo;DefaultBypassCheck.IsChecked=options.DefaultToBypass;
        SizeCombo.SelectedItem=EnvelopePreset.All.FirstOrDefault(p=>p.Key==options.Preset)??EnvelopePreset.All[0];
        WidthText.Text=options.WidthMm.ToString("0.###",CultureInfo.InvariantCulture);HeightText.Text=options.HeightMm.ToString("0.###",CultureInfo.InvariantCulture);
        WidthText.IsEnabled=HeightText.IsEnabled=!options.IsLabel&&options.Preset=="Custom";
        FontText.Text=options.FontPt.ToString(CultureInfo.InvariantCulture);ReturnFontText.Text=options.ReturnFontPt.ToString(CultureInfo.InvariantCulture);
        OffsetXText.Text=options.OffsetXmm.ToString(CultureInfo.InvariantCulture);OffsetYText.Text=options.OffsetYmm.ToString(CultureInfo.InvariantCulture);
        CopiesText.Text=options.Copies.ToString(CultureInfo.InvariantCulture);RotateCheck.IsChecked=options.Rotate180;
        desiredPrinter=options.Printer;desiredTrayId=options.TrayId;desiredTrayName=options.TrayName;
        if(PrinterCombo.Items.Contains(options.Printer))PrinterCombo.SelectedItem=options.Printer;
        if(printerInfo?.Name==options.Printer)TrayCombo.SelectedItem=printerInfo.Trays.FirstOrDefault(t=>t.Id==options.TrayId&&t.Name.Equals(options.TrayName,StringComparison.OrdinalIgnoreCase));
        if((historySource?ReturnCombo.Items.OfType<ReturnChoice>().LastOrDefault(c=>(c.Address?.Id??"")==options.ReturnAddressId):ReturnCombo.Items.OfType<ReturnChoice>().FirstOrDefault(c=>(c.Address?.Id??"")==options.ReturnAddressId)) is {} choice)ReturnCombo.SelectedItem=choice;
        loading=previous;imageable=null;
    }
    private void LoadReturnChoices(MailingAddress? historical)
    {
        var choices=new List<ReturnChoice>{new("No return address",null)};
        choices.AddRange(book.Offices.Select(a=>new ReturnChoice(a.DisplayName,a)));
        if(historical is not null)choices.Add(new("Historical return address — "+historical.DisplayName,historical));
        ReturnCombo.ItemsSource=choices;ReturnCombo.SelectedItem=historical is null?choices[0]:choices[^1];
    }
    private void LoadRows(IEnumerable<MailingAddress> addresses,IEnumerable<string> checkedIds)
    {
        var selected=checkedIds.ToHashSet(StringComparer.Ordinal);rows=addresses.Select(a=>new RecipientRow(a){Selected=selected.Contains(a.Id)}).ToList();
        foreach(var row in rows)row.PropertyChanged+=(_,e)=>{if(!loading&&e.PropertyName==nameof(RecipientRow.Selected)){UpdateCounts();UpdatePreview();}};
        recipientsView=CollectionViewSource.GetDefaultView(rows);recipientsView.Filter=item=>item is RecipientRow row&&row.Address.Text.Contains(SearchText.Text.Trim(),StringComparison.OrdinalIgnoreCase);
        RecipientsGrid.ItemsSource=recipientsView;RecipientsGrid.SelectedItem=recipientsView.Cast<RecipientRow>().FirstOrDefault();
        SourceText.Text=historySource?"Historical addresses are shown as saved in the original request.":$"{rows.Count} contractor addresses in your saved address book.";
        CurrentBookButton.Visibility=historySource?Visibility.Visible:Visibility.Collapsed;UpdateCounts();
    }
    private static readonly MailingAddress PreviewPlaceholder = new()
    {
        Id="preview-placeholder", Company="[Contractors]", Street="[Mailing Address]",
        City="[City]", Province="[Province]", Postal="[Postal Code]"
    };
    private MailingAddress? CurrentAddress=>RecipientsGrid.SelectedItem is RecipientRow { Selected:true } row
        ? row.Address : rows.FirstOrDefault(r=>r.Selected)?.Address;
    private void UpdateCounts()
    {
        int selected=rows.Count(r=>r.Selected);string copies=CopiesText.Text;
        SelectionText.Text=$"{selected} recipients selected · {copies} copies each";
        if(Labels&&IncludeReturnLabels.IsChecked==true)SelectionText.Text+=" - "+ReturnLabelCountText.Text+" return-address labels";
    }
    private void UpdatePreview()
    {
        LoadingInstructions.Text=Labels?"Load a full Letter label sheet (8.5 x 11 inches), even for one label position. The printer tray must be registered as Letter / Labels.":$"Envelope: {WidthText.Text} x {HeightText.Text} mm. Create the Word document, then choose the printer and Bypass Tray in Word's Printer Properties.";

        if(loading)return;
        try
        {
            UpdateReturnLogoControls();UpdateCounts();
            var options=ReadOptions();var sender=(ReturnCombo.SelectedItem as ReturnChoice)?.Address;
            var drawing=CurrentDrawing(true);PreviewImage.Source=new DrawingImage(drawing.Drawing);
            PreviewCaption.Text=CurrentCaption()+" - scaled to fit";
            PreviewWarnings.Text=string.Join(" ",drawing.Warnings);
            GuideText.Text=Labels?"Position numbers and dashed lines are preview guides only. They are never printed.":"Dashed lines are preview guides only. Print the generated envelope document from Word.";
        }catch(Exception ex){PreviewWarnings.Text=ex.Message;PreviewImage.Source=null;}
    }
    private void UpdateReturnLogoControls()
    {
        bool available=(ReturnCombo.SelectedItem as ReturnChoice)?.Address?.LogoPng is not null;
        IncludeReturnLogo.IsEnabled=available;
        if(!available){bool previous=loading;loading=true;IncludeReturnLogo.IsChecked=false;loading=previous;}
        ReturnLogoHint.Text=available?"Uses this return address's saved logo. Preview shows its printed position.":"Add a logo under Manage addresses - Return addresses.";
    }
    private void DropdownWheel(object s,MouseWheelEventArgs e)
    {if(s is ComboBox {IsDropDownOpen:true})return;e.Handled=true;EnvelopeScroll.ScrollToVerticalOffset(EnvelopeScroll.VerticalOffset-e.Delta/120.0*48);}
    private void Layout_Changed(object s,RoutedEventArgs e){if(loading)return;imageable=null;UpdatePreview();}
    private void Field_Changed(object s,TextChangedEventArgs e){if(loading)return;imageable=null;UpdatePreview();}
    private void Size_Changed(object s,SelectionChangedEventArgs e)
    {
        if(loading||Labels||SizeCombo.SelectedItem is not EnvelopePreset preset)return;
        loading=true;WidthText.IsEnabled=HeightText.IsEnabled=preset.Key=="Custom";
        if(preset.Key!="Custom"){WidthText.Text=preset.WidthMm.ToString(CultureInfo.InvariantCulture);HeightText.Text=preset.HeightMm.ToString(CultureInfo.InvariantCulture);}
        loading=false;LoadProfile();imageable=null;UpdatePreview();
    }
    private void LoadProfile()
    {
        try
        {
            string name=PrinterCombo.SelectedItem as string??"",size=Labels?"Label-"+SelectedLabel.Key:(SizeCombo.SelectedItem as EnvelopePreset)?.Key??"10";
            var profile=store.LoadProfile(name,size);if(profile is null)return;
            profile.ReturnAddressId=(ReturnCombo.SelectedItem as ReturnChoice)?.Address?.Id??"";ApplyOptions(profile);
            PrinterStatus.Text="Saved settings loaded for this printer and mailing size.";
        }catch(Exception ex){PrinterStatus.Text="Saved profile could not be loaded: "+ex.Message;}
    }
    private async Task RefreshPrintersAsync(bool useProfile)
    {
        try
        {
            string current=PrinterCombo.SelectedItem as string??desiredPrinter;
            var names=await printer.PrintersAsync();loading=true;PrinterCombo.ItemsSource=names;
            PrinterCombo.SelectedItem=names.Contains(current)?current:names.FirstOrDefault(n=>n.Contains("SHARP",StringComparison.OrdinalIgnoreCase)&&n.Contains("BP-70C31",StringComparison.OrdinalIgnoreCase));loading=false;
            if(PrinterCombo.SelectedItem is null){TrayCombo.ItemsSource=null;PrinterStatus.Text="The office printer is not installed here. Preview and history remain available; choose the Sharp printer at the office.";return;}
            await LoadPrinterAsync(useProfile);
        }catch(Exception ex){loading=false;PrinterStatus.Text="Could not read Windows printers. "+ex.Message;}
    }
    private async Task LoadPrinterAsync(bool useProfile)
    {
        int revision=++printerRevision;string? name=PrinterCombo.SelectedItem as string;if(name is null)return;
        try
        {
            imageable=null;if(printer is IEnvelopePrinterProperties properties)properties.ClearDriverPreferences();
            if(useProfile){desiredTrayId=null;desiredTrayName="";LoadProfile();}
            var info=await printer.DescribeAsync(name);if(revision!=printerRevision||PrinterCombo.SelectedItem as string!=name)return;
            printerInfo=info;desiredPrinter=name;loading=true;TrayCombo.ItemsSource=info.Trays;
            TrayCombo.SelectedItem=PrinterSessionChoices.InitialTray(info.Trays,new(){TrayId=desiredTrayId,TrayName=desiredTrayName,DefaultToBypass=DefaultBypassCheck.IsChecked==true},info.DefaultTrayId);
            loading=false;PrinterStatus.Text=info.Driver+" · "+info.Notes;
            if(TrayCombo.SelectedItem is null)PrinterStatus.Text+=" Select an available paper tray; no substitute has been selected.";
            UpdatePreview();
        }catch(Exception ex){loading=false;PrinterStatus.Text="Printer information could not be read. "+ex.Message;}
    }
    private async void Printer_Changed(object s,SelectionChangedEventArgs e){if(!loading&&!busy)await LoadPrinterAsync(true);}
    private void Tray_Changed(object s,SelectionChangedEventArgs e){if(loading)return;if(printer is IEnvelopePrinterProperties properties)properties.ClearDriverPreferences();imageable=null;UpdatePreview();}
    private void DefaultBypass_Changed(object s,RoutedEventArgs e)
    {
        if(loading)return;
        if(printer is IEnvelopePrinterProperties properties)properties.ClearDriverPreferences();
        if(printerInfo is not null){loading=true;TrayCombo.SelectedItem=PrinterSessionChoices.InitialTray(printerInfo.Trays,new(){DefaultToBypass=DefaultBypassCheck.IsChecked==true},printerInfo.DefaultTrayId);loading=false;}
        imageable=null;UpdatePreview();
    }
    private async void RefreshPrinters_Click(object s,RoutedEventArgs e){if(!busy)await RefreshPrintersAsync(false);}
    private void Search_Changed(object s,TextChangedEventArgs e)
    {if(loading)return;recipientsView?.Refresh();if(RecipientsGrid.SelectedItem is null)RecipientsGrid.SelectedItem=recipientsView?.Cast<RecipientRow>().FirstOrDefault();UpdatePreview();}
    private void Recipient_Changed(object s,SelectionChangedEventArgs e)=>UpdatePreview();
    private void SelectShown_Click(object s,RoutedEventArgs e){loading=true;foreach(var row in recipientsView?.Cast<RecipientRow>()??[])row.Selected=true;loading=false;UpdateCounts();UpdatePreview();}
    private void ClearSelection_Click(object s,RoutedEventArgs e){loading=true;foreach(var row in rows)row.Selected=false;loading=false;UpdateCounts();UpdatePreview();}
    private void Move(int direction)
    {if(Labels){var snapshot=Snapshot(preview:true);labelPage=(labelPage+direction+snapshot.Pages)%snapshot.Pages;UpdatePreview();return;}var visible=rows.Where(r=>r.Selected).ToList();if(visible.Count==0)return;int index=visible.FindIndex(r=>r.Address==CurrentAddress);if(index<0)index=0;RecipientsGrid.SelectedItem=visible[(index+direction+visible.Count)%visible.Count];RecipientsGrid.ScrollIntoView(RecipientsGrid.SelectedItem);}
    private void Previous_Click(object s,RoutedEventArgs e)=>Move(-1);
    private void Next_Click(object s,RoutedEventArgs e)=>Move(1);
    private EnvelopeSnapshot Snapshot(bool one=false,bool preview=false)
    {
        var options=ReadOptions();
        bool returnOnly=Labels&&options.ReturnLabelCount>0&&!rows.Any(r=>r.Selected);
        if(one){options.Copies=1;options.ReturnLabelCount=returnOnly?1:0;options.IncludeReturnLogo=options.IncludeReturnLogo&&(!Labels||options.ReturnLabelCount>0);options.LayoutVersion=options.ExpectedLayoutVersion;}
        var recipients=one&&!returnOnly?new List<MailingAddress>{CurrentAddress??throw new InvalidOperationException("Select at least one contractor first.")}:one?[]:rows.Where(r=>r.Selected).Select(r=>r.Address).ToList();
        if(preview&&recipients.Count==0&&options.ReturnLabelCount==0)recipients.Add(PreviewPlaceholder);
        var snapshot=new EnvelopeSnapshot(recipients,Labels&&options.ReturnLabelCount==0?null:(ReturnCombo.SelectedItem as ReturnChoice)?.Address,options);snapshot.Validate();return snapshot;
    }
    private void EnlargePreview_Click(object s,RoutedEventArgs e)
    {
        try{new EnvelopePreviewWindow(CurrentDrawing(true),CurrentCaption(),Labels){Owner=OwnerWindow}.ShowDialog();}
        catch(Exception ex){Error(ex);}
    }
    private void SaveProfile_Click(object s,RoutedEventArgs e){try{store.SaveProfile(ReadOptions());PrinterStatus.Text="Alignment and tray choices saved. Printer properties changes apply to the current app session.";}catch(Exception ex){Error(ex);}}
    private void SetBusy(bool value)
    {
        busy=value;UpdateHistoryActions();SettingsPanel.IsEnabled=!value;RecipientsPanel.IsEnabled=PreviewActions.IsEnabled=!value;PrintButton.IsEnabled=PrinterPropertiesButton.IsEnabled=!value;CreateWordButton.IsEnabled=!value;CancelPrintButton.IsEnabled=value;
    }
    private async void ValidatePrinter_Click(object s,RoutedEventArgs e)
    {
        if(busy||!Labels)return;
        try{SetBusy(true);using var prepared=await printer.PrepareAsync(Snapshot(one:true));imageable=prepared.Imageable;PrinterStatus.Text="Validated: "+prepared.Snapshot.Options.TrayName+" · "+prepared.MediaType+" · sides as selected in Printer properties. No job submitted.";UpdatePreview();}
        catch(OperationCanceledException){PrinterStatus.Text="Printer setup cancelled. No job submitted.";}
        catch(Exception ex){Error(ex);}finally{SetBusy(false);}
    }
    private void DriverPreferences_Click(object s,RoutedEventArgs e)
    {
        if(busy||!Labels)return;
        try{ConfigureDriverPreferences();}catch(Exception ex){Error(ex);}
    }
    private bool ConfigureDriverPreferences()
    {
        if(printer is not IEnvelopePrinterProperties properties)throw new InvalidOperationException("Printer properties are unavailable for this printer.");
        var options=ReadOptions();
        if(!properties.DriverPreferences(options,new WindowInteropHelper(OwnerWindow).Handle))return false;
        if(properties.SelectedTray is {} selected&&printerInfo is not null)
        {
            loading=true;TrayCombo.SelectedItem=printerInfo.Trays.FirstOrDefault(t=>t.Id==selected.Id&&t.Name==selected.Name);loading=false;
        }
        imageable=null;PrinterStatus.Text="Printer properties applied for this printer and layout in the current app session. Print will use these settings; Windows defaults were not changed.";UpdatePreview();
        return true;
    }
    private async void Diagnostics_Click(object s,RoutedEventArgs e)
    {
        try
        {
            string name=PrinterCombo.SelectedItem as string??throw new InvalidOperationException("Select a printer first.");var info=await printer.DescribeAsync(name);
            string text="Printer: "+info.Name+"\nDriver: "+info.Driver+"\n"+info.Notes+"\n\nPaper trays:\n"+string.Join("\n",info.Trays.Select(t=>$"{t.Id}: {t.Name}"))+"\n\nDriver paper sizes:\n"+string.Join("\n",info.Papers.Select(p=>$"{p.Name}: {p.WidthMm:0.#} × {p.HeightMm:0.#} mm (ID {p.Id})"))+"\n\nRead-only diagnostics. No print job was submitted.";
            if(printer is WindowsEnvelopePrinter windows)text+="\n\nLast layout check:\n"+windows.LastLayoutDiagnostics;
            ShowText("Printer diagnostics",text);
        }catch(Exception ex){Error(ex);}
    }
    private async Task PrintAsync()
    {
        if(busy||!Labels)return;
        try
        {
            var snapshot=Snapshot();store.SaveOptions(ReadOptions());SetBusy(true);cancellation=new();PrintProgress.Value=0;PrintProgress.IsIndeterminate=true;
            StatusText.Text="Validating the selected printer, paper size, and tray…";
            var entry=await new EnvelopePrintCoordinator(store,printer).PrintAsync(snapshot,parentHistory is null?"Print":"Reprint",parentHistory,prepared=>
            {
                cancellation.Token.ThrowIfCancellationRequested();imageable=prepared.Imageable;UpdatePreview();
                var o=prepared.Snapshot.Options;
                string labelDetails=o.IsLabel?"\nTemplate: "+o.LabelSheet!.Label+"\nReturn-address labels: "+o.ReturnLabelCount+"\nFirst-sheet slots used: "+string.Join(", ",MailingRenderer.Placements(prepared.Snapshot,0).Select(p=>p.Slot)):"";
                bool yes=MessageBox.Show(OwnerWindow,$"Print {prepared.Snapshot.Pages} page(s), {prepared.Snapshot.Items} {(o.IsLabel?"label(s)":"envelope(s)")}?\n\nPrinter: {o.Printer}\nTray: {o.TrayName}\nSize: {o.WidthMm:0.#} × {o.HeightMm:0.#} mm\nMedia: {prepared.MediaType}{labelDetails}\nSides: as selected in Printer properties\nTo: {string.Join("; ",(prepared.Snapshot.Recipients.Select(a=>a.Company).Concat(o.ReturnLabelCount>0?new[]{"Return address: "+prepared.Snapshot.ReturnAddress!.Company}:Array.Empty<string>())).Take(3))}{(prepared.Snapshot.Recipients.Count>3?" ...":"")}\n\nLoad the matching {(o.IsLabel?"label sheets":"envelopes")} and confirm the tray's size/type on the printer.","Confirm mailing print",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)==MessageBoxResult.Yes;
                return Task.FromResult(yes);
            },p=>Dispatcher.Invoke(()=>{PrintProgress.IsIndeterminate=false;PrintProgress.Value=p.TotalPages>0?p.PagesRendered*100.0/p.TotalPages:0;StatusText.Text=$"{p.PagesRendered} of {p.TotalPages} pages sent to Windows. Physical printing is not confirmed.";}),cancellation.Token);
            StatusText.Text=entry.Status+". "+entry.Events.Last().Details;
        }
        catch(Exception ex){StatusText.Text="Request did not finish. Check request history before reprinting. "+ex.Message;Error(ex);}
        finally{PrintProgress.IsIndeterminate=false;cancellation?.Dispose();cancellation=null;SetBusy(false);RefreshHistory();}
    }
    private async void Print_Click(object s,RoutedEventArgs e)=>await PrintAsync();
    private void CancelPrint_Click(object s,RoutedEventArgs e){Cancel();CancelPrintButton.IsEnabled=false;StatusText.Text=Labels?"Cancellation requested. Pages already handed to the printer may still print; check the printer before retrying.":"Cancelling Word document generation.";}
    private void ManageAddresses_Click(object s,RoutedEventArgs e)
    {
        if(busy)return;if(addressBookError is not null){Error(new InvalidDataException(addressBookError));return;}
        var editor=new AddressBookWindow(book){Owner=OwnerWindow};if(editor.ShowDialog()!=true)return;
        try
        {
            store.SaveBook(editor.Book);book=editor.Book;
            if(!historySource)
            {
                var selected=rows.Where(r=>r.Selected).Select(r=>r.Address.Id).ToList();string returnId=(ReturnCombo.SelectedItem as ReturnChoice)?.Address?.Id??"";
                loading=true;LoadReturnChoices(null);ReturnCombo.SelectedItem=ReturnCombo.Items.OfType<ReturnChoice>().FirstOrDefault(c=>(c.Address?.Id??"")==returnId)??ReturnCombo.Items[0];LoadRows(book.Contractors,selected);loading=false;UpdatePreview();
            }
            StatusText.Text="Address book saved. Existing history snapshots are unchanged.";
        }catch(Exception ex){Error(ex);}
    }
    private void CurrentBook_Click(object s,RoutedEventArgs e)
    {historySource=false;parentHistory=null;loading=true;LoadReturnChoices(null);SearchText.Clear();LoadRows(book.Contractors,[]);loading=false;UpdatePreview();}
    private void PrepareTab_Click(object s,RoutedEventArgs e)=>ShowHistory(false);
    private void HistoryTab_Click(object s,RoutedEventArgs e){if(busy)return;if(HistoryPanel.Visibility==Visibility.Visible){ShowHistory(false);return;}RefreshHistory();ShowHistory(true);}
    private void ShowHistory(bool show)
    {
        EnvelopeScroll.Visibility=show?Visibility.Collapsed:Visibility.Visible;HistoryPanel.Visibility=show?Visibility.Visible:Visibility.Collapsed;
        HistoryNavigationButton.Content=show?"Back to feature":"View history";
    }
    private void RefreshHistory()
    {
        try{var selected=HistoryGrid.SelectedItems.Cast<EnvelopeHistoryEntry>().Select(e=>e.Id).ToHashSet();history=store.History(out int unreadable);FilterHistory();foreach(var row in HistoryGrid.Items.Cast<EnvelopeHistoryEntry>().Where(e=>selected.Contains(e.Id)))HistoryGrid.SelectedItems.Add(row);HistoryStatus.Text=$"{history.Count} requests retained. Times shown in Example City time."+(unreadable>0?$" {unreadable} unreadable history file(s) were not shown.":"");}
        catch(Exception ex){HistoryStatus.Text="History could not be read. "+ex.Message;}
    }
    private void FilterHistory()
    {
        if(HistoryGrid is null)return;string query=HistorySearch.Text.Trim();DateTime? from=FromDate.SelectedDate,to=ToDate.SelectedDate;
        HistoryGrid.ItemsSource=history.Where(entry=>
        {var day=TimeZoneInfo.ConvertTime(entry.RequestedUtc,TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")).Date;
            return(!from.HasValue||day>=from.Value.Date)&&(!to.HasValue||day<=to.Value.Date)&&
                (query.Length==0||entry.Id.Contains(query,StringComparison.OrdinalIgnoreCase)||entry.Status.Contains(query,StringComparison.OrdinalIgnoreCase)||entry.Snapshot.Options.Printer.Contains(query,StringComparison.OrdinalIgnoreCase)||entry.Snapshot.Recipients.Any(a=>a.Text.Contains(query,StringComparison.OrdinalIgnoreCase))||entry.Snapshot.ReturnAddress?.Text.Contains(query,StringComparison.OrdinalIgnoreCase)==true);}).ToList();
    }
    private void HistorySearch_Changed(object s,TextChangedEventArgs e){if(!loading)FilterHistory();}
    private void HistoryDate_Changed(object s,SelectionChangedEventArgs e){if(!loading)FilterHistory();}
    private void RefreshHistory_Click(object s,RoutedEventArgs e)=>RefreshHistory();
    private void UpdateHistoryActions()
    {
        int count=HistoryGrid.SelectedItems.Count;
        DeleteHistoryButton.IsEnabled=!busy&&count>0;
        DeleteHistoryButton.Content=count>0?$"Delete selected history ({count})":"Delete selected history";
        HistoryDetailsButton.IsEnabled=LoadHistoryButton.IsEnabled=HistoryQueueButton.IsEnabled=!busy&&count==1;
    }
    private void HistorySelection_Changed(object s,SelectionChangedEventArgs e)=>UpdateHistoryActions();
    private void DeleteHistory_Click(object s,RoutedEventArgs e)
    {
        if(busy)return;
        try
        {
            DeleteSelectedHistory(entries=>MessageBox.Show(OwnerWindow,
                $"Delete {entries.Count} selected saved history record(s)?\n\nThis cannot be undone. It removes only the saved records and does not cancel print jobs or change saved addresses.",
                "Delete mailing history",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No)==MessageBoxResult.Yes);
        }
        catch(Exception ex){Error(ex);}
    }
    internal bool DeleteSelectedHistory(Func<IReadOnlyList<EnvelopeHistoryEntry>,bool> confirm)
    {
        if(busy)return false;
        var selected=HistoryGrid.SelectedItems.Cast<EnvelopeHistoryEntry>().ToList();
        if(selected.Count==0)throw new InvalidOperationException("Select at least one history entry first.");
        if(!confirm(selected))return false;
        var result=store.DeleteHistoryBatch(selected.Select(e=>e.Id));
        if(parentHistory is not null&&result.Deleted.Contains(parentHistory))parentHistory=null;
        RefreshHistory();UpdateHistoryActions();
        HistoryStatus.Text=$"{result.Deleted.Count} selected history record(s) deleted. "+(result.Failed.Count>0?$"{result.Failed.Count} could not be deleted; they remain selected for retry. ":"")+HistoryStatus.Text;
        return result.Deleted.Count>0;
    }
    private EnvelopeHistoryEntry SelectedHistory()=>HistoryGrid.SelectedItems.Count==1?HistoryGrid.SelectedItems[0] as EnvelopeHistoryEntry??throw new InvalidOperationException("Select a history entry first."):throw new InvalidOperationException("Select exactly one history entry for this action.");
    private void HistoryDetails_Click(object s,RoutedEventArgs e)
    {
        try{var entry=SelectedHistory();ShowText("Envelope request "+entry.Id,
            $"Requested (Example City): {entry.RequestedLocal}\nWindows user: {entry.WindowsUser}\nType: {entry.Kind}\nStatus: {entry.Status}\nPrinter: {entry.Snapshot.Options.Printer}\nTray: {entry.Snapshot.Options.TrayName} (ID {entry.Snapshot.Options.TrayId})\nLayout: {entry.EnvelopeSize}\nCopies per recipient: {entry.Snapshot.Options.Copies}\nPages: {entry.Pages}\nMailing items: {entry.Snapshot.Items}\nReturn-address labels: {entry.Snapshot.Options.ReturnLabelCount}\nFirst label slot: {entry.Snapshot.Options.FirstLabelSlot}\nSpecific first-sheet slots: {(entry.Snapshot.Options.FirstPageSlots is {} slots?string.Join(", ",slots):"Use start slot")}\nWindows job ID: {entry.JobId?.ToString()??"Unavailable"}\nParent request: {entry.ParentId??"None"}\n\nReturn address:\n{entry.Snapshot.ReturnAddress?.Text??"None"}\n\nRecipients:\n"+string.Join("\n\n",entry.Snapshot.Recipients.Select(a=>a.Text))+"\n\nEvents:\n"+string.Join("\n",entry.Events.Select(ev=>$"{TimeZoneInfo.ConvertTime(ev.WhenUtc,TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")):yyyy-MM-dd HH:mm:ss zzz} · {ev.Status}\n{ev.Details}")));}
        catch(Exception ex){Error(ex);}
    }
    private void ShowText(string title,string text)
    {
        var box=new TextBox{Text=text,IsReadOnly=true,TextWrapping=TextWrapping.Wrap,VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Padding=new Thickness(16),BorderThickness=new Thickness(0)};
        var window=new Window{Title=title,Owner=OwnerWindow,Width=760,Height=600,MinWidth=600,MinHeight=400,WindowStartupLocation=WindowStartupLocation.CenterOwner,Content=box};window.ShowDialog();
    }
    internal async Task LoadHistoryAsync(EnvelopeHistoryEntry entry)
    {
        if(busy)throw new InvalidOperationException("Finish the current request before loading history.");
        entry.Snapshot.Validate();historySource=true;parentHistory=entry.Id;loading=true;SearchText.Clear();LoadReturnChoices(entry.Snapshot.ReturnAddress);
        ApplyOptions(entry.Snapshot.Options);if(entry.Snapshot.ReturnAddress is not null)ReturnCombo.SelectedIndex=ReturnCombo.Items.Count-1;
        LoadRows(entry.Snapshot.Recipients,entry.Snapshot.Recipients.Select(a=>a.Id));loading=false;ShowHistory(false);UpdatePreview();await RefreshPrintersAsync(false);
        StatusText.Text="Historical request loaded using its saved addresses and layout. Nothing has been printed. Printing again creates a linked new request.";
    }
    private async void LoadHistory_Click(object s,RoutedEventArgs e){try{await LoadHistoryAsync(SelectedHistory());}catch(Exception ex){Error(ex);}}
    private async void QueueStatus_Click(object s,RoutedEventArgs e)
    {
        try{var entry=SelectedHistory();if(entry.JobId is not int id)throw new InvalidOperationException("No Windows job ID was captured for this request. Its physical outcome cannot be inferred.");
            string status=await printer.QueueStatusAsync(entry.Snapshot.Options.Printer,id,entry.DocumentName);store.Update(entry.Id,status,"Read from Windows queue. A missing job is not proof of physical printing.");RefreshHistory();}
        catch(Exception ex){Error(ex);}
    }
    internal static string Csv(IEnumerable<EnvelopeHistoryEntry> entries)
    {
        static string Cell(string value)=>"\""+((value.Length>0&&"=+-@\t\r".Contains(value[0]))?"'":"")+value.Replace("\"","\"\"")+"\"";
        var csv=new StringBuilder("Request ID,Requested Example City,Type,Status,Pages,Recipients,Printer,Tray,Size,Job ID,Parent request,Labels or envelopes,First label slot,Specific first-sheet slots,Return-address labels\r\n");
        foreach(var e in entries)csv.AppendLine(string.Join(",",new[]{e.Id,e.RequestedLocal,e.Kind,e.Status,e.Pages.ToString(),string.Join("; ",e.Snapshot.Recipients.Select(a=>a.Company)),e.Snapshot.Options.Printer,e.Snapshot.Options.TrayName,e.EnvelopeSize,e.JobId?.ToString()??"",e.ParentId??"",e.Snapshot.Items.ToString(),e.Snapshot.Options.IsLabel?e.Snapshot.Options.FirstLabelSlot.ToString():"",e.Snapshot.Options.FirstPageSlots is {} slots?string.Join("; ",slots):"",e.Snapshot.Options.ReturnLabelCount.ToString()}.Select(Cell)));
        return csv.ToString();
    }
    private void ExportHistory_Click(object s,RoutedEventArgs e)
    {
        try{var picker=new SaveFileDialog{Title="Export envelope history",Filter="CSV (*.csv)|*.csv",DefaultExt=".csv",FileName="Envelope-history.csv",OverwritePrompt=true};if(picker.ShowDialog(OwnerWindow)==true)File.WriteAllText(picker.FileName,Csv(HistoryGrid.Items.Cast<EnvelopeHistoryEntry>()),new UTF8Encoding(true));}
        catch(Exception ex){Error(ex);}
    }
}
