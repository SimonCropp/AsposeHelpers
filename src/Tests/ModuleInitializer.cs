public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Initialize()
    {
        VerifyAspose.Initialize();
        VerifyDiffPlex.Initialize();
        VerifyImageMagick.RegisterComparers(.5);
        VerifierSettings.IgnoreMember("Width");
        VerifierSettings.ScrubLinesContaining("evaluation", "License");
    }
}