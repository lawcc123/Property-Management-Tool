using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Data.Sqlite;

namespace PropertyManager.App.Operations;

internal sealed record OperationsRun(string Id,string Started,string Action,string Status,string Details)
{
    public string LocalTime=>TimeZoneInfo.ConvertTime(DateTimeOffset.Parse(Started),TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")).ToString("yyyy-MM-dd HH:mm:ss");
}
internal sealed class OperationsHistory
{
    private readonly string path;
    internal OperationsHistory(string? directory=null)
    {
        directory??=OperationsSettings.DirectoryPath;Directory.CreateDirectory(directory);path=Path.Combine(directory,"history.db");
        using var db=Open();using var c=db.CreateCommand();c.CommandText="CREATE TABLE IF NOT EXISTS runs(id TEXT PRIMARY KEY,started TEXT,action TEXT,status TEXT,details BLOB,send_token TEXT UNIQUE,schedule_token TEXT UNIQUE);";c.ExecuteNonQuery();
    }
    private SqliteConnection Open(){var db=new SqliteConnection(new SqliteConnectionStringBuilder{DataSource=path,Pooling=false}.ToString());db.Open();using var c=db.CreateCommand();c.CommandText="PRAGMA secure_delete=ON; PRAGMA synchronous=FULL; PRAGMA journal_mode=DELETE;";c.ExecuteNonQuery();return db;}
    private static byte[] Protect(string text)=>ProtectedData.Protect(Encoding.UTF8.GetBytes(text),Encoding.UTF8.GetBytes("PropertyManagerDemo.Operations.v1"),DataProtectionScope.CurrentUser);
    private static string Reveal(byte[] bytes)=>Encoding.UTF8.GetString(ProtectedData.Unprotect(bytes,Encoding.UTF8.GetBytes("PropertyManagerDemo.Operations.v1"),DataProtectionScope.CurrentUser));
    internal string Start(string action,string? sendToken=null,string? scheduleToken=null)
    {
        string id=Guid.NewGuid().ToString("N");using var db=Open();using var c=db.CreateCommand();c.CommandText="INSERT INTO runs VALUES($id,$time,$action,'Running',$details,$send,$schedule)";
        c.Parameters.AddWithValue("$id",id);c.Parameters.AddWithValue("$time",DateTimeOffset.UtcNow.ToString("O"));c.Parameters.AddWithValue("$action",action);c.Parameters.AddWithValue("$details",Protect("Preparing reminder"));c.Parameters.AddWithValue("$send",(object?)sendToken??DBNull.Value);c.Parameters.AddWithValue("$schedule",(object?)scheduleToken??DBNull.Value);
        try{c.ExecuteNonQuery();}catch(SqliteException ex)when(ex.SqliteErrorCode==19){throw new InvalidOperationException("This preview was already used to send, or today's scheduled reminder already started. Refresh and review before sending manually.");}return id;
    }
    internal void Update(string id,string status,string details)
    {
        using var db=Open();using var c=db.CreateCommand();c.CommandText="UPDATE runs SET status=$status,details=$details WHERE id=$id";c.Parameters.AddWithValue("$id",id);c.Parameters.AddWithValue("$status",status);c.Parameters.AddWithValue("$details",Protect(details));c.ExecuteNonQuery();
    }
    internal List<OperationsRun> Read()
    {
        using var db=Open();using var c=db.CreateCommand();c.CommandText="SELECT id,started,action,status,details FROM runs ORDER BY started DESC LIMIT 200";using var r=c.ExecuteReader();List<OperationsRun> rows=[];while(r.Read())rows.Add(new(r.GetString(0),r.GetString(1),r.GetString(2),r.GetString(3),Reveal((byte[])r[4])));return rows;
    }
}
