using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using PropertyManager.App.Services;
namespace PropertyManager.App.Settings;

// Phase 1: shared settings/address data. Recovery journals and sending states are
// deliberately kept in their existing stores until their own migration is tested.
internal static class SettingsDatabase
{
    internal static string Root=>AppSettings.DataDirectory;
    internal static string DatabasePath=>Path.Combine(Root,"app.db");
    internal static readonly string[] Files=["settings.json","appearance-settings.json","buildinglink-settings.json","OperationsReminder/settings.json","OutlookSync/quick-sync.json","OutlookSync/group-names.json","EnvelopePrinter/address-book.json","EnvelopePrinter/last-options.json","EnvelopePrinter/printer-profiles.json"];
    private static readonly byte[] Entropy=Encoding.UTF8.GetBytes("PropertyManagerDemo.SharedSettings.v1");
    private static string Hash(string text)=>Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
    private static string? Key(string path)
    {
        string relative=Path.GetRelativePath(Root,Path.GetFullPath(path)).Replace('\\','/');
        return Files.FirstOrDefault(f=>f.Equals(relative,StringComparison.OrdinalIgnoreCase));
    }
    private static SqliteConnection Open(string path,bool readOnly=false)
    {
        var db=new SqliteConnection(new SqliteConnectionStringBuilder{DataSource=path,Mode=readOnly?SqliteOpenMode.ReadOnly:SqliteOpenMode.ReadWriteCreate,Pooling=false}.ToString());db.Open();
        if(!readOnly){using var c=db.CreateCommand();c.CommandText="PRAGMA synchronous=FULL; PRAGMA secure_delete=ON;";c.ExecuteNonQuery();}return db;
    }
    internal static bool Active
    {
        get
        {
            if(!File.Exists(DatabasePath))return false;
            using var db=Open(DatabasePath,true);using var c=db.CreateCommand();c.CommandText="SELECT version FROM schema_version WHERE id=1";
            long version=Convert.ToInt64(c.ExecuteScalar());if(version!=1)throw new InvalidDataException("This settings database requires a different app version. Existing files were kept.");return true;
        }
    }
    private static string Decode(byte[] bytes,string expected)
    {
        string text=Encoding.UTF8.GetString(ProtectedData.Unprotect(bytes,Entropy,DataProtectionScope.CurrentUser));
        if(Hash(text)!=expected)throw new InvalidDataException("Settings verification failed. Restore a verified backup before continuing.");using var json=JsonDocument.Parse(text);return text;
    }
    internal static string? Read(string path)
    {
        string? key=Key(path);if(key is null||!Active)return File.Exists(path)?File.ReadAllText(path):null;
        using var db=Open(DatabasePath,true);using var c=db.CreateCommand();c.CommandText="SELECT payload,digest FROM settings WHERE name=$name";c.Parameters.AddWithValue("$name",key);
        using var reader=c.ExecuteReader();if(!reader.Read())return null;
        return Decode((byte[])reader[0],reader.GetString(1));
    }
    internal static bool Exists(string path)=>Read(path) is not null;
    private static void Insert(SqliteConnection db,string key,string text,SqliteTransaction? transaction=null)
    {
        using var c=db.CreateCommand();c.Transaction=transaction;c.CommandText="INSERT INTO settings VALUES($name,$payload,$digest,$updated) ON CONFLICT(name) DO UPDATE SET payload=excluded.payload,digest=excluded.digest,updated=excluded.updated";
        c.Parameters.AddWithValue("$name",key);c.Parameters.AddWithValue("$payload",ProtectedData.Protect(Encoding.UTF8.GetBytes(text),Entropy,DataProtectionScope.CurrentUser));c.Parameters.AddWithValue("$digest",Hash(text));c.Parameters.AddWithValue("$updated",DateTimeOffset.UtcNow.ToString("O"));c.ExecuteNonQuery();
    }
    internal static void AtomicFile(string path,string text)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);string stage=path+"."+Guid.NewGuid().ToString("N")+".tmp";
        try{File.WriteAllText(stage,text);File.Move(stage,path,true);}finally{if(File.Exists(stage))File.Delete(stage);}
    }
    internal static void Write(string path,string text,bool onlyIfMissing=false)
    {
        using var json=JsonDocument.Parse(text);
        if(Key(path) is null){if(!onlyIfMissing||!File.Exists(path))AtomicFile(path,text);return;}
        Directory.CreateDirectory(Root);
        using var held=new FileStream(Path.Combine(Root,"settings-write.lock"),FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);
        if(onlyIfMissing&&Exists(path))return;
        string? key=Key(path);if(key is null||!Active){AtomicFile(path,text);return;}
        using var db=Open(DatabasePath);using var tx=db.BeginTransaction();Insert(db,key,text,tx);
        string? old=File.Exists(path)?File.ReadAllText(path):null;
        try{AtomicFile(path,text);tx.Commit();}
        catch{try{if(old is null)File.Delete(path);else AtomicFile(path,old);}catch{}throw;}
    }
    internal static string Verify()
    {
        if(!Active)return "Shared settings have not been migrated. Existing files are still in use.";
        using var db=Open(DatabasePath,true);using var check=db.CreateCommand();check.CommandText="PRAGMA integrity_check";if((string?)check.ExecuteScalar()!="ok")throw new InvalidDataException("Database integrity check failed.");
        using var c=db.CreateCommand();c.CommandText="SELECT payload,digest FROM settings";using var r=c.ExecuteReader();int count=0;while(r.Read()){Decode((byte[])r[0],r.GetString(1));count++;}
        return $"Verified: database integrity, decryption and content checksums passed for {count} settings/address records. Recovery, sending history and drafts remain in their existing stores.";
    }
    private static void ValidateLegacy(string key,string text)
    {
        using var json=JsonDocument.Parse(text);if(json.RootElement.ValueKind!=JsonValueKind.Object)throw new InvalidDataException(key+": expected a settings object.");
        switch(key)
        {
            case "settings.json": JsonSerializer.Deserialize<AppSettings>(text);break;
            case "appearance-settings.json": (JsonSerializer.Deserialize<Appearance.AppearanceSettings>(text)??throw new InvalidDataException("Appearance settings are empty.")).Validate();break;
            case "buildinglink-settings.json": JsonSerializer.Deserialize<BuildingLink.BuildingLinkSettings>(text);break;
            case "OperationsReminder/settings.json": JsonSerializer.Deserialize<Operations.OperationsSettings>(text);break;
            case "OutlookSync/quick-sync.json": JsonSerializer.Deserialize<OutlookSync.QuickSyncSettings>(text);break;
            case "OutlookSync/group-names.json": foreach(var names in (JsonSerializer.Deserialize<OutlookSync.GroupNaming>(text)??throw new InvalidDataException("Group settings are empty.")).Corporations.Values)OutlookSync.GroupNaming.Validate(names);break;
            case "EnvelopePrinter/address-book.json": (JsonSerializer.Deserialize<Envelopes.AddressBook>(text)??throw new InvalidDataException("Address book is empty.")).Validate();break;
            case "EnvelopePrinter/last-options.json": (JsonSerializer.Deserialize<Envelopes.EnvelopeOptions>(text)??throw new InvalidDataException("Printer settings are empty.")).CurrentEnvelopeSettings().Validate();break;
            case "EnvelopePrinter/printer-profiles.json": foreach(var value in (JsonSerializer.Deserialize<Dictionary<string,Envelopes.EnvelopeOptions>>(text)??throw new InvalidDataException("Printer profiles are empty.")).Values)value.CurrentEnvelopeSettings().Validate();break;
        }
    }
    internal static string Migrate(string backupDestination,Action<string>? beforeVerify=null)
    {
        Directory.CreateDirectory(Root);using var held=new FileStream(Path.Combine(Root,"settings-write.lock"),FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);
        using var operations=new SettingsOperationLocks(settings:false);
        if(Active)return Verify();
        string backup=AppDataBackup.CreateWhileLocked(backupDestination);
        var originals=new Dictionary<string,string>();
        string stage=DatabasePath+"."+Guid.NewGuid().ToString("N")+".tmp";
        try
        {
            foreach(string key in Files)
            {
                string source=Path.Combine(Root,key);if(!File.Exists(source))continue;
                if((File.GetAttributes(source)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("A settings source is a link. Review its location before migration.");
                string text=File.ReadAllText(source);ValidateLegacy(key,text);originals[key]=text;
                string copy=Path.Combine(backup,key);
                if(!File.Exists(copy)||File.ReadAllText(copy)!=text)throw new InvalidDataException("The verified app-data backup did not match its settings source. Retry after all operations have stopped.");
            }
            using(var db=Open(stage))
            {
                using(var schema=db.CreateCommand()){schema.CommandText="CREATE TABLE schema_version(id INTEGER PRIMARY KEY CHECK(id=1),version INTEGER NOT NULL);INSERT INTO schema_version VALUES(1,1);CREATE TABLE settings(name TEXT PRIMARY KEY,payload BLOB NOT NULL,digest TEXT NOT NULL,updated TEXT NOT NULL);CREATE TABLE migrations(id TEXT PRIMARY KEY,completed TEXT NOT NULL,records INTEGER NOT NULL,backup TEXT NOT NULL);";schema.ExecuteNonQuery();}
                using(var tx=db.BeginTransaction()){foreach(var (key,text) in originals)Insert(db,key,text,tx);tx.Commit();}
            }
            beforeVerify?.Invoke(stage);
            using(var db=Open(stage,true))
            {
                using var check=db.CreateCommand();check.CommandText="PRAGMA integrity_check";if((string?)check.ExecuteScalar()!="ok")throw new InvalidDataException("Migration integrity check failed.");
                using var c=db.CreateCommand();c.CommandText="SELECT name,payload,digest FROM settings";using var records=c.ExecuteReader();int count=0;
                while(records.Read()){string key=records.GetString(0);if(!originals.TryGetValue(key,out string? old)||Decode((byte[])records[1],records.GetString(2))!=old)throw new InvalidDataException("Migration content comparison failed.");count++;}
                if(count!=originals.Count)throw new InvalidDataException("Migration record count does not match.");
            }
            foreach(var (key,text) in originals)if(File.ReadAllText(Path.Combine(Root,key))!=text)throw new InvalidOperationException("Settings changed during migration. Nothing was activated; retry when all operations have stopped.");
            foreach(string key in Files.Except(originals.Keys))if(File.Exists(Path.Combine(Root,key)))throw new InvalidOperationException("New settings appeared during migration. Nothing was activated; retry.");
            using(var db=Open(stage)){using var c=db.CreateCommand();c.CommandText="INSERT INTO migrations VALUES($id,$time,$count,$backup)";c.Parameters.AddWithValue("$id",Guid.NewGuid().ToString("N"));c.Parameters.AddWithValue("$time",DateTimeOffset.UtcNow.ToString("O"));c.Parameters.AddWithValue("$count",originals.Count);c.Parameters.AddWithValue("$backup",backup);c.ExecuteNonQuery();}
            string report=JsonSerializer.Serialize(new{Version=1,Verified=true,Records=originals.Count,Checks=new[]{"backup content comparison","database integrity","Windows decryption","exact content comparison","source unchanged"},Unmigrated=new[]{"Outlook recovery","sending history and duplicate-send protection","reminder history","drafts","signatures","credentials","browser sessions"}},new JsonSerializerOptions{WriteIndented=true});
            File.WriteAllText(Path.Combine(backup,"verification.json"),report);
            File.Move(stage,DatabasePath,false);return Verify()+"\nOriginal files retained. Backup and report: "+backup;
        }
        catch{File.WriteAllText(Path.Combine(backup,"migration-failed.txt"),"Migration did not activate. Original settings remain in place. Resolve the error shown in the app before retrying.");throw;}
        finally{if(File.Exists(stage))File.Delete(stage);}
    }
}
