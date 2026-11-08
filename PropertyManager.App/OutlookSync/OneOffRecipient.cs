using System.IO;
using System.Text;
using PropertyManager.App.EmailLists;

namespace PropertyManager.App.OutlookSync;

// MS-OXCDATA 2.2.5.1 / MAPI one-off EntryID. No directory or contact writes.
internal static class OneOffRecipient
{
    internal static string EntryId(GroupMember member)
    {
        if(!GeneratedEmailListService.IsValidEmail(member.Email)||
            string.IsNullOrWhiteSpace(member.Name)||member.Name.Contains('\0')||member.Email.Contains('\0'))
            throw new InvalidDataException("A group member needs a valid email and display name.");
        using var stream=new MemoryStream();using var writer=new BinaryWriter(stream,Encoding.Unicode,true);
        writer.Write(0u);
        writer.Write(Convert.FromHexString("812B1FA4BEA310199D6E00DD010F5402"));
        writer.Write((ushort)0); // Version
        writer.Write((ushort)0x9001); // Unicode, no address-book lookup, no TNEF
        foreach(string value in new[]{member.Name,"SMTP",member.Email.ToLowerInvariant()})
        {writer.Write(Encoding.Unicode.GetBytes(value));writer.Write((ushort)0);}
        return Convert.ToHexString(stream.ToArray());
    }
}
