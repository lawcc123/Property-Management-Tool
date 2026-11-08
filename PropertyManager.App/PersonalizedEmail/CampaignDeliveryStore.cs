using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Data.Sqlite;
namespace PropertyManager.App.PersonalizedEmail;
internal sealed class CampaignDeliveryStore
{
    internal string DirectoryPath {get;}
    internal string DatabasePath=>Path.Combine(DirectoryPath,"delivery.db");
    private static readonly byte[] Entropy=Encoding.UTF8.GetBytes("PropertyManagerDemo.CampaignDelivery.v1");
    internal CampaignDeliveryStore(string? directory=null)
    {
        DirectoryPath=directory??Path.Combine(Services.AppSettings.DataDirectory,"PersonalizedEmail","Sending");Directory.CreateDirectory(DirectoryPath);
        using var db=Open();using var cmd=db.CreateCommand();cmd.CommandText="PRAGMA journal_mode=DELETE; CREATE TABLE IF NOT EXISTS campaigns(id TEXT PRIMARY KEY,created TEXT NOT NULL,status TEXT NOT NULL,fingerprint TEXT NOT NULL,mode TEXT NOT NULL,data BLOB NOT NULL); CREATE TABLE IF NOT EXISTS messages(id TEXT PRIMARY KEY,campaign TEXT NOT NULL REFERENCES campaigns(id),state TEXT NOT NULL,detail BLOB NOT NULL); CREATE TABLE IF NOT EXISTS campaign_cleanup(campaign TEXT PRIMARY KEY REFERENCES campaigns(id),completed TEXT NOT NULL);";cmd.ExecuteNonQuery();
    }
    private SqliteConnection Open(){var db=new SqliteConnection(new SqliteConnectionStringBuilder{DataSource=DatabasePath,Pooling=false,ForeignKeys=true}.ToString());db.Open();using var c=db.CreateCommand();c.CommandText="PRAGMA synchronous=FULL; PRAGMA secure_delete=ON;";c.ExecuteNonQuery();return db;}
    private static byte[] Pack<T>(T value)=>ProtectedData.Protect(JsonSerializer.SerializeToUtf8Bytes(value),Entropy,DataProtectionScope.CurrentUser);
    private static T Read<T>(byte[] data)=>JsonSerializer.Deserialize<T>(ProtectedData.Unprotect(data,Entropy,DataProtectionScope.CurrentUser))!;
    internal FileStream Lock()=>new(Path.Combine(DirectoryPath,"sending.lock"),FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);
    internal void RecoverInterrupted()
    {
        try{using var held=Lock();using var db=Open();using var c=db.CreateCommand();c.CommandText="UPDATE messages SET state='Uncertain' WHERE state='Submitting'; UPDATE campaigns SET status='Interrupted - review history' WHERE status='Sending';";c.ExecuteNonQuery();foreach(var campaign in PendingCleanup())CompletedCampaignFiles.TryCleanup(campaign,this);}catch(IOException){}
    }
    internal PreparedCampaign? ExistingLive(string fingerprint)
    {
        using var db=Open();using var c=db.CreateCommand();c.CommandText="SELECT data FROM campaigns WHERE fingerprint=$fingerprint AND mode='Live' ORDER BY created DESC LIMIT 1";c.Parameters.AddWithValue("$fingerprint",fingerprint);var data=c.ExecuteScalar();return data is byte[] bytes?Read<PreparedCampaign>(bytes):null;
    }
    internal void Add(PreparedCampaign campaign)
    {
        using var db=Open();using var tx=db.BeginTransaction();using var c=db.CreateCommand();c.Transaction=tx;c.CommandText="INSERT INTO campaigns VALUES($id,$created,'Prepared',$fingerprint,$mode,$data)";c.Parameters.AddWithValue("$id",campaign.Id);c.Parameters.AddWithValue("$created",DateTimeOffset.UtcNow.ToString("O"));c.Parameters.AddWithValue("$fingerprint",campaign.Fingerprint);c.Parameters.AddWithValue("$mode",campaign.Mode);c.Parameters.AddWithValue("$data",Pack(campaign));c.ExecuteNonQuery();
        foreach(var message in campaign.Messages){using var m=db.CreateCommand();m.Transaction=tx;m.CommandText="INSERT INTO messages VALUES($id,$campaign,'Not attempted',$detail)";m.Parameters.AddWithValue("$id",message.Id);m.Parameters.AddWithValue("$campaign",campaign.Id);m.Parameters.AddWithValue("$detail",Pack(""));m.ExecuteNonQuery();}tx.Commit();
    }
    internal List<CampaignRun> Runs()
    {
        using var db=Open();using var c=db.CreateCommand();c.CommandText="SELECT id,created,status,data FROM campaigns ORDER BY created DESC";using var reader=c.ExecuteReader();var rows=new List<CampaignRun>();while(reader.Read())rows.Add(new(reader.GetString(0),reader.GetString(1),reader.GetString(2),Read<PreparedCampaign>((byte[])reader[3])));return rows;
    }
    internal List<DeliveryRow> Rows(PreparedCampaign campaign)
    {
        using var db=Open();using var c=db.CreateCommand();c.CommandText="SELECT id,state,detail FROM messages WHERE campaign=$id";c.Parameters.AddWithValue("$id",campaign.Id);using var reader=c.ExecuteReader();var values=new Dictionary<string,(string State,string Detail)>();while(reader.Read())values[reader.GetString(0)]=(reader.GetString(1),Read<string>((byte[])reader[2]));return campaign.Messages.Select(m=>new DeliveryRow(m,values[m.Id].State,values[m.Id].Detail)).ToList();
    }
    internal void SetMessage(string id,string state,string detail)
    {
        using var db=Open();using var c=db.CreateCommand();c.CommandText="UPDATE messages SET state=$state,detail=$detail WHERE id=$id";c.Parameters.AddWithValue("$id",id);c.Parameters.AddWithValue("$state",state);c.Parameters.AddWithValue("$detail",Pack(detail));if(c.ExecuteNonQuery()!=1)throw new InvalidOperationException("The sending history record was not found.");
    }
    internal bool FilesCleaned(string id)
    {
        using var db=Open();using var c=db.CreateCommand();c.CommandText="SELECT 1 FROM campaign_cleanup WHERE campaign=$id";c.Parameters.AddWithValue("$id",id);return c.ExecuteScalar() is not null;
    }
    private List<PreparedCampaign> PendingCleanup()
    {
        using var db=Open();using var c=db.CreateCommand();c.CommandText="SELECT c.data FROM campaigns c WHERE c.mode='Live' AND NOT EXISTS(SELECT 1 FROM campaign_cleanup k WHERE k.campaign=c.id) AND EXISTS(SELECT 1 FROM messages m WHERE m.campaign=c.id) AND NOT EXISTS(SELECT 1 FROM messages m WHERE m.campaign=c.id AND m.state<>'Accepted')";
        using var reader=c.ExecuteReader();var result=new List<PreparedCampaign>();while(reader.Read())result.Add(Read<PreparedCampaign>((byte[])reader[0]));return result;
    }
    internal void MarkFilesCleaned(PreparedCampaign campaign)
    {
        var rows=Rows(campaign);using var db=Open();using var tx=db.BeginTransaction();
        foreach(var row in rows)
        {
            using var m=db.CreateCommand();m.Transaction=tx;m.CommandText="UPDATE messages SET detail=$detail WHERE id=$id";m.Parameters.AddWithValue("$id",row.Message.Id);m.Parameters.AddWithValue("$detail",Pack(row.Detail+" Temporary delivery attachments were removed after the completed live send. Saved Step 3 exports are kept."));m.ExecuteNonQuery();
        }
        using var c=db.CreateCommand();c.Transaction=tx;c.CommandText="INSERT OR IGNORE INTO campaign_cleanup VALUES($id,$date)";c.Parameters.AddWithValue("$id",campaign.Id);c.Parameters.AddWithValue("$date",DateTimeOffset.UtcNow.ToString("O"));c.ExecuteNonQuery();tx.Commit();
    }
    internal void SetRun(string id,string status){using var db=Open();using var c=db.CreateCommand();c.CommandText="UPDATE campaigns SET status=$status WHERE id=$id";c.Parameters.AddWithValue("$id",id);c.Parameters.AddWithValue("$status",status);c.ExecuteNonQuery();}
}
