using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Data.Sqlite;

namespace PropertyManager.App.OutlookSync;

internal sealed record RunContext(OutlookMailbox Mailbox,string Corporation,string Mode,string? SourcePath=null,string? SourceHash=null,List<string>? Corporations=null);
internal sealed record RecoveryRun(string Id,string Started,string Status,RunContext Context)
{
    public string LocalTime=>TimeZoneInfo.ConvertTime(DateTimeOffset.Parse(Started),TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")).ToString("yyyy-MM-dd HH:mm:ss");
    public string Corporation=>Context.Corporation;
    public string Mailbox=>Context.Mailbox.Label;
    public string Mode=>Context.Mode;
}
internal sealed record RecoveryOperation(long Id,OutlookEntity? Before,OutlookEntity? Desired,OutlookEntity? Actual,string Status);
internal sealed class RecoveryStore
{
    internal string DirectoryPath {get;}
    internal string DatabasePath=>Path.Combine(DirectoryPath,"outlook-recovery.db");
    internal RecoveryStore(string? directory=null)
    {
        DirectoryPath=directory??Path.Combine(Services.AppSettings.DataDirectory,"OutlookSync");Directory.CreateDirectory(DirectoryPath);
        using var db=Open();using var cmd=db.CreateCommand();
        cmd.CommandText="""
        PRAGMA journal_mode=DELETE;
        CREATE TABLE IF NOT EXISTS runs(id TEXT PRIMARY KEY,started TEXT NOT NULL,status TEXT NOT NULL,scope TEXT NOT NULL,context BLOB NOT NULL);
        CREATE TABLE IF NOT EXISTS operations(id INTEGER PRIMARY KEY AUTOINCREMENT,run_id TEXT NOT NULL REFERENCES runs(id) ON DELETE CASCADE,before BLOB,desired BLOB,actual BLOB,status TEXT NOT NULL);
        CREATE INDEX IF NOT EXISTS run_operations ON operations(run_id,id);
        CREATE TABLE IF NOT EXISTS recovery_info(key TEXT PRIMARY KEY,value TEXT NOT NULL);
        INSERT OR IGNORE INTO recovery_info VALUES('legacy_history',CASE WHEN EXISTS(SELECT 1 FROM runs) THEN '1' ELSE '0' END);
        """;cmd.ExecuteNonQuery();
    }
    private SqliteConnection Open()
    {
        var db=new SqliteConnection(new SqliteConnectionStringBuilder{DataSource=DatabasePath,Pooling=false,ForeignKeys=true}.ToString());db.Open();
        using var cmd=db.CreateCommand();cmd.CommandText="PRAGMA secure_delete=ON; PRAGMA synchronous=FULL;";cmd.ExecuteNonQuery();return db;
    }
    private static byte[] Encrypt<T>(T value)=>ProtectedData.Protect(JsonSerializer.SerializeToUtf8Bytes(value),Encoding.UTF8.GetBytes("PropertyManagerDemo.OutlookSync.v1"),DataProtectionScope.CurrentUser);
    private static T Decrypt<T>(byte[] data)=>JsonSerializer.Deserialize<T>(ProtectedData.Unprotect(data,Encoding.UTF8.GetBytes("PropertyManagerDemo.OutlookSync.v1"),DataProtectionScope.CurrentUser))!;
    private static object Pack(OutlookEntity? value)=>value is null?DBNull.Value:Encrypt(value);
    private static OutlookEntity? Unpack(object value)=>value is DBNull?null:Decrypt<OutlookEntity>((byte[])value);
    internal string Start(SyncPlan plan,List<string>? corporations=null)
    {
        string id=Guid.NewGuid().ToString("N");using var db=Open();using var cmd=db.CreateCommand();
        cmd.CommandText="INSERT INTO runs VALUES($id,$time,'Applying',$scope,$context)";
        cmd.Parameters.AddWithValue("$id",id);cmd.Parameters.AddWithValue("$time",DateTimeOffset.UtcNow.ToString("O"));
        cmd.Parameters.AddWithValue("$scope",Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(plan.Mailbox.StoreId+"|"+plan.Corporation))));
        cmd.Parameters.AddWithValue("$context",Encrypt(new RunContext(plan.Mailbox,plan.Corporation,plan.Mode,plan.SourcePath,plan.SourceHash,corporations)));cmd.ExecuteNonQuery();return id;
    }
    internal long Prepare(string run,OutlookEntity? before,OutlookEntity? desired)
    {
        using var db=Open();using var cmd=db.CreateCommand();cmd.CommandText="INSERT INTO operations(run_id,before,desired,status) VALUES($run,$before,$desired,'Pending'); SELECT last_insert_rowid();";
        cmd.Parameters.AddWithValue("$run",run);cmd.Parameters.AddWithValue("$before",Pack(before));cmd.Parameters.AddWithValue("$desired",Pack(desired));return (long)cmd.ExecuteScalar()!;
    }
    internal void Applied(long operation,OutlookEntity? actual)
    {
        using var db=Open();using var cmd=db.CreateCommand();cmd.CommandText="UPDATE operations SET actual=$actual,status='Applied' WHERE id=$id";
        cmd.Parameters.AddWithValue("$actual",Pack(actual));cmd.Parameters.AddWithValue("$id",operation);cmd.ExecuteNonQuery();
    }
    internal void OperationStatus(long operation,string status)
    {
        using var db=Open();using var cmd=db.CreateCommand();cmd.CommandText="UPDATE operations SET status=$status WHERE id=$id";cmd.Parameters.AddWithValue("$id",operation);cmd.Parameters.AddWithValue("$status",status);cmd.ExecuteNonQuery();
    }
    internal void Finish(string run,string status)
    {
        using var db=Open();using var cmd=db.CreateCommand();cmd.CommandText="UPDATE runs SET status=$status WHERE id=$id";cmd.Parameters.AddWithValue("$id",run);cmd.Parameters.AddWithValue("$status",status);cmd.ExecuteNonQuery();
    }
    internal List<RecoveryRun> Runs()
    {
        using var db=Open();using var cmd=db.CreateCommand();cmd.CommandText="SELECT id,started,status,context FROM runs ORDER BY started DESC";using var reader=cmd.ExecuteReader();var result=new List<RecoveryRun>();
        while(reader.Read())result.Add(new(reader.GetString(0),reader.GetString(1),reader.GetString(2),Decrypt<RunContext>((byte[])reader[3])));return result;
    }
    internal List<RecoveryOperation> Operations(string run)
    {
        using var db=Open();using var cmd=db.CreateCommand();cmd.CommandText="SELECT id,before,desired,actual,status FROM operations WHERE run_id=$run ORDER BY id DESC";cmd.Parameters.AddWithValue("$run",run);using var reader=cmd.ExecuteReader();var result=new List<RecoveryOperation>();
        while(reader.Read())result.Add(new(reader.GetInt64(0),Unpack(reader[1]),Unpack(reader[2]),Unpack(reader[3]),reader.GetString(4)));return result;
    }
    internal bool HasLegacyHistory
    {
        get{using var db=Open();using var cmd=db.CreateCommand();cmd.CommandText="SELECT value FROM recovery_info WHERE key='legacy_history'";return (string?)cmd.ExecuteScalar()=="1";}
    }
    internal void Retain()
    {
        // Full undo needs every item's first before-image and its complete write chain.
        // Recovery history is no longer automatically pruned.
    }
    internal List<(RecoveryRun Run,RecoveryOperation Operation)> Journal(OutlookMailbox mailbox)
        =>Runs().Where(r=>r.Context.Mailbox.StoreId==mailbox.StoreId&&r.Context.Mailbox.FolderId==mailbox.FolderId)
            .SelectMany(r=>Operations(r.Id).Select(o=>(Run:r,Operation:o))).OrderBy(v=>v.Operation.Id).ToList();
}
