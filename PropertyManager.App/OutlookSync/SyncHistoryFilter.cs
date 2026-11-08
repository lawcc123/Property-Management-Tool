namespace PropertyManager.App.OutlookSync;

internal static class SyncHistoryFilter
{
    internal static List<RecoveryRun> Matching(IEnumerable<RecoveryRun> runs,DateTime? from,DateTime? through)
    {
        var zone=TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
        return runs.Select(run=>(Run:run,Started:DateTimeOffset.Parse(run.Started)))
            .Where(x=>
            {
                DateTime day=TimeZoneInfo.ConvertTime(x.Started,zone).Date;
                return (!from.HasValue||day>=from.Value.Date)&&(!through.HasValue||day<=through.Value.Date);
            }).OrderByDescending(x=>x.Started).ThenByDescending(x=>x.Run.Id).Select(x=>x.Run).ToList();
    }
}
