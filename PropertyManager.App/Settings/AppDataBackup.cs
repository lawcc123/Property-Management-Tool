using System.IO;
using System.Text.Json;
using Microsoft.Data.Sqlite;
namespace PropertyManager.App.Settings;
internal static class AppDataBackup
{
    // A same-Windows-account recovery copy, not a portable credential export.
    internal static string Create(string destination)
    {
        using var held=new SettingsOperationLocks();
        return CreateWhileLocked(destination);
    }
    // Caller must hold both the shared-settings and feature-operation locks.
    internal static string CreateWhileLocked(string destination)
    {
        if(string.IsNullOrWhiteSpace(destination))throw new ArgumentException("Choose a backup folder.");
        string root=Path.GetFullPath(SettingsDatabase.Root),chosen=Path.GetFullPath(destination);
        if(chosen.Equals(root,StringComparison.OrdinalIgnoreCase)||chosen.StartsWith(root+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("Choose a backup folder outside the app-data folder.");
        string target=Path.Combine(chosen,"PropertyManagerDemo-backup-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N")[..8]);Directory.CreateDirectory(target);
        var files=new List<string>();
        try
        {
            var standalone=SettingsDatabase.Files.Concat(new[]{"app.db","OutlookSync/outlook-recovery.db","OutlookSync/last-error.dpapi","OperationsReminder/history.db","PersonalizedEmail/Sending/delivery.db","PersonalizedEmail/signature.dpapi","PersonalizedEmail/signature-library.dpapi","occupant-folder-shortcuts.json"});
            var paths=standalone.Select(f=>Path.Combine(root,f)).Where(File.Exists).ToList();
            foreach(string folder in new[]{"EnvelopePrinter/History","EnvelopePrinter/DriverProfiles","PersonalizedEmail/Drafts","PersonalizedEmail/SentDrafts"})
            {
                string source=Path.Combine(root,folder);if(!Directory.Exists(source))continue;
                var directories=new Stack<string>();directories.Push(source);
                while(directories.Count>0)
                {
                    string current=directories.Pop();if((File.GetAttributes(current)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("A backup source is a link. Review it before backing up.");
                    paths.AddRange(Directory.GetFiles(current));foreach(string child in Directory.GetDirectories(current))directories.Push(child);
                }
            }
            foreach(string path in paths.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if((File.GetAttributes(path)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("A backup source is a link.");
                string relative=Path.GetRelativePath(root,path),copy=Path.Combine(target,relative);Directory.CreateDirectory(Path.GetDirectoryName(copy)!);
                if(Path.GetExtension(path)==".db")
                {
                    using var source=new SqliteConnection(new SqliteConnectionStringBuilder{DataSource=path,Mode=SqliteOpenMode.ReadOnly,Pooling=false}.ToString());source.Open();
                    using var output=new SqliteConnection(new SqliteConnectionStringBuilder{DataSource=copy,Pooling=false}.ToString());output.Open();source.BackupDatabase(output);
                    using var check=output.CreateCommand();check.CommandText="PRAGMA integrity_check";if((string?)check.ExecuteScalar()!="ok")throw new InvalidDataException("A database backup failed integrity verification.");
                }
                else
                {
                    File.Copy(path,copy);if(!System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(path)).SequenceEqual(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(copy))))throw new IOException("A backup file changed while being copied. Retry when operations have stopped.");
                }
                files.Add(relative);
            }
            File.WriteAllText(Path.Combine(target,"backup-report.json"),JsonSerializer.Serialize(new{Complete=true,Files=files,SameWindowsAccountRequired=true,Excluded=new[]{"passwords and browser sessions","external Excel workbooks and exported documents","document caches and temporary delivery attachments","Windows scheduled tasks"}},new JsonSerializerOptions{WriteIndented=true}));return target;
        }
        catch{File.WriteAllText(Path.Combine(target,"INCOMPLETE.txt"),"This backup did not complete. Do not use it for recovery; retry after resolving the error.");throw;}
    }
}
