namespace PropertyManager.App;

internal static class BuildFeatures
{
#if DEVELOPER_TOOLS
    internal static bool DeveloperToolsEnabled => true;
#else
    internal static bool DeveloperToolsEnabled => false;
#endif
}
