public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Initialize()
    {
        VerifyAspose.Initialize();
        VerifyDiffPlex.Initialize();
        VerifyImageMagick.RegisterComparers(.01);
        VerifierSettings.IgnoreMember("Width");
        VerifierSettings.ScrubLinesContaining("evaluation", "License");
    }
}