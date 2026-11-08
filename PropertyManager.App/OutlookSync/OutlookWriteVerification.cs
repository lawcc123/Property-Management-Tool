namespace PropertyManager.App.OutlookSync;

internal static class OutlookWriteVerification
{
    internal static List<string> Differences(OutlookEntity? desired,OutlookEntity? actual)
    {
        if(desired is null)return actual is null?[]:["item was not removed"];
        if(actual is null)return ["saved item could not be read"];
        var expected=desired;
        if(expected.SameContent(actual))return [];
        var differences=expected.ChangedFields(actual).Where(f=>f!="modification timestamp only").ToList();
        return differences.Count>0?differences:["saved sync state differs"];
    }
}
