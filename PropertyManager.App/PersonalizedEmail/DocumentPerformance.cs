using System.Diagnostics;
using System.IO;
using System.Text.Json;
namespace PropertyManager.App.PersonalizedEmail;

// Small bounded diagnostics: timings/counts only, never names, paths or contents.
internal static class DocumentPerformance
{
    private static readonly object gate=new();
    internal static string LogPath=>Path.Combine(Services.AppSettings.DataDirectory,"PersonalizedEmail","Diagnostics","document-performance.jsonl");
#if DEVELOPER_TOOLS
    internal static IDisposable Measure(string stage)=>new Measurement(stage);
    internal static void Record(string stage,double elapsedMs=0,long? bytes=null)
    {
        try
        {
            lock(gate)
            {
                string path=LogPath;Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                if(File.Exists(path)&&new FileInfo(path).Length>1_000_000)File.Move(path,path+".previous",true);
                File.AppendAllText(path,JsonSerializer.Serialize(new{Utc=DateTimeOffset.UtcNow,Stage=stage,ElapsedMs=Math.Round(elapsedMs,2),Bytes=bytes})+Environment.NewLine);
            }
        }
        catch(IOException){}catch(UnauthorizedAccessException){}
    }
    private sealed class Measurement(string stage):IDisposable
    {
        private readonly Stopwatch watch=Stopwatch.StartNew();
        public void Dispose(){watch.Stop();Record(stage,watch.Elapsed.TotalMilliseconds);}
    }
#else
    internal static IDisposable Measure(string stage)=>NoTiming.Instance;
    internal static void Record(string stage,double elapsedMs=0,long? bytes=null) { }
    private sealed class NoTiming:IDisposable
    {
        internal static readonly NoTiming Instance=new();
        public void Dispose() { }
    }
#endif

}
