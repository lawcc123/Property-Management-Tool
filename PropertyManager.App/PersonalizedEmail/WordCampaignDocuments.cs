using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
namespace PropertyManager.App.PersonalizedEmail;
internal sealed class WordCampaignDocuments:ICampaignDocuments
{
    private object? application;
    internal WordCampaignDocuments()
    {
        using var timing=DocumentPerformance.Measure("word.startup");
        var type=Type.GetTypeFromProgID("Word.Application")??throw new InvalidOperationException("Microsoft Word is required for personalized Word/PDF documents. Install desktop Word, or send email with existing common attachments.");
        application=Activator.CreateInstance(type)??throw new InvalidOperationException("Microsoft Word could not start.");
        dynamic word=application;word.Visible=false;word.DisplayAlerts=0;word.AutomationSecurity=3;
    }
    private static void Release(object? value){if(value is not null&&Marshal.IsComObject(value))Marshal.ReleaseComObject(value);}
    public string Generate(string template,string output,bool personalize,IReadOnlyDictionary<string,string> fields)
    {
        using var timing=DocumentPerformance.Measure("word.generate.total");
        object? document=null;string? intermediate=null;
        try
        {
            dynamic word=application!;
            using(DocumentPerformance.Measure("word.open"))document=word.Documents.Open(template,ConfirmConversions:false,ReadOnly:true,AddToRecentFiles:false,Visible:false,OpenAndRepair:false);
            dynamic doc=document;
            if(personalize)
            {
                using var mergeTiming=DocumentPerformance.Measure("word.personalize");
                foreach(dynamic root in doc.StoryRanges)
                {
                    dynamic? story=root;
                    while(story is not null)
                    {
                        dynamic native=story.Fields;
                        for(int i=(int)native.Count;i>=1;i--)
                        {
                            dynamic field=native.Item(i);string code=(string)field.Code.Text;
                            var match=Regex.Match(code,@"MERGEFIELD\s+(?:""(?<name>[^""]+)""|(?<name>[^\s\\]+))",RegexOptions.IgnoreCase);
                            if(match.Success)
                            {
                                string key=match.Groups["name"].Value;if(!fields.TryGetValue(key,out string? value))throw new InvalidDataException("Missing document field: "+key);
                                field.Result.Text=value;field.Unlink();
                            }
                            Release(field);
                        }
                        string text=(string)story.Text;int start=(int)story.Start;
                        foreach(Match token in Regex.Matches(text,@"\{\{(?<name>[^{}]+)\}\}").Cast<Match>().Reverse())
                        {
                            string key=token.Groups["name"].Value.Trim();if(!fields.TryGetValue(key,out string? value))throw new InvalidDataException("Missing document field: "+key);
                            dynamic range=story.Duplicate;range.SetRange(start+token.Index,start+token.Index+token.Length);range.Text=value;Release(range);
                        }
                        dynamic? next=story.NextStoryRange;Release(native);Release(story);story=next;
                    }
                }
            }
            bool pdf=Path.GetExtension(output).Equals(".pdf",StringComparison.OrdinalIgnoreCase);
            string check=pdf?(intermediate=output+"."+Guid.NewGuid().ToString("N")+".docx"):output;
            using(DocumentPerformance.Measure("word.save.docx"))doc.SaveAs2(check,16,AddToRecentFiles:false);
            if(pdf)using(DocumentPerformance.Measure("word.export.pdf"))doc.ExportAsFixedFormat(output,17,OpenAfterExport:false,IncludeDocProps:false);
            using(DocumentPerformance.Measure("word.close")){doc.Close(0);Release(document);document=null;}
            // Word holds an exclusive file lock until Close. Validate the saved
            // document only after releasing it, before any email can be sent.
            using var checkTiming=DocumentPerformance.Measure("word.validate.fields");
            if(personalize&&CampaignService.DocumentFields(check).Count>0)throw new InvalidDataException("The document still contains unfilled personalization fields. Check headers, text boxes and template fields before sending.");
            return output;
        }
        finally
        {
            if(document is not null){try{((dynamic)document).Close(0);}finally{Release(document);}}
            if(intermediate is not null){try{File.Delete(intermediate);}catch(IOException){}}
        }
    }
    public void Dispose(){if(application is null)return;using var timing=DocumentPerformance.Measure("word.shutdown");try{((dynamic)application).Quit(0);}finally{Release(application);application=null;}}
}
