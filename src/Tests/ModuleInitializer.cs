using AsposeHelpers;

public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Initialize()
    {
        ApplyAsposeLicense();
        VerifyAspose.Initialize();
        VerifyDiffPlex.Initialize();
        VerifyImageMagick.RegisterComparers(.6);
        VerifierSettings.IgnoreMember("Width");
        VerifierSettings.ScrubLinesContaining("evaluation", "License");
    }

    static void ApplyAsposeLicense()
    {
        var licenseText = Environment.GetEnvironmentVariable("AsposeLicense");
        if (licenseText == null)
        {
            throw new("Expected a `AsposeLicense` environment variable");
        }

        var stream = new MemoryStream();
        var writer = new StreamWriter(stream);
        writer.Write(licenseText);
        writer.Flush();

        AsposeLicense.Apply(stream);
    }
}