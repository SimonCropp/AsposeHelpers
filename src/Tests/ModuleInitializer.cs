public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Initialize()
    {
        VerifyAspose.Initialize();
        VerifyDiffPlex.Initialize();
        VerifyImageMagick.RegisterComparers(.003);
        VerifierSettings.IgnoreMember("Width");
    }
}