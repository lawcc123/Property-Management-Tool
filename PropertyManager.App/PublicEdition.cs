namespace PropertyManager.App;
internal static class PublicEdition
{
    internal static void BlockLiveOperation() => throw new System.InvalidOperationException(
        "This public showcase does not connect to live services, create scheduled tasks or submit print jobs.");
}
