namespace PropertyManager.App.OutlookSync;

internal sealed record GroupMembershipDelta(List<GroupMember> Remove,List<GroupMember> Add)
{
    internal static GroupMember Key(GroupMember member)=>member with{Email=member.Email.Trim().ToLowerInvariant()};
    internal static GroupMembershipDelta Build(IReadOnlyList<GroupMember> before,IReadOnlyList<GroupMember> after)
    {
        // Compare a multiset: retain unchanged native recipients, remove only
        // excess copies, and replace a label by removing/adding that member.
        var remaining=after.GroupBy(Key).ToDictionary(g=>g.Key,g=>g.Count());
        var remove=new List<GroupMember>();var add=new List<GroupMember>();
        foreach(var member in before)
        {
            var key=Key(member);
            if(remaining.GetValueOrDefault(key)>0)remaining[key]--;
            else remove.Add(member);
        }
        foreach(var member in after)
        {
            var key=Key(member);
            if(remaining.GetValueOrDefault(key)>0){add.Add(member);remaining[key]--;}
        }
        return new(remove,add);
    }
}
