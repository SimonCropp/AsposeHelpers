namespace AsposeHelpers;

public static class AsposeLicense
{
    public static void Apply(Stream stream)
    {
        Email(stream);
        Pdf(stream);
        Cells(stream);
        Word(stream);
        Slides(stream);
    }

    static void Slides(Stream licenseStream)
    {
        var lic = new Aspose.Slides.License();
        licenseStream.Position = 0;
        lic.SetLicense(licenseStream);
    }

    static void Word(Stream licenseStream)
    {
        var lic = new Aspose.Words.License();
        licenseStream.Position = 0;
        lic.SetLicense(licenseStream);
    }

    static void Email(Stream licenseStream)
    {
        var lic = new Aspose.Email.License();
        licenseStream.Position = 0;
        lic.SetLicense(licenseStream);
    }

    static void Pdf(Stream licenseStream)
    {
        var lic = new Aspose.Pdf.License();
        licenseStream.Position = 0;
        lic.SetLicense(licenseStream);
    }

    static void Cells(Stream licenseStream)
    {
        var lic = new Aspose.Cells.License();
        licenseStream.Position = 0;
        lic.SetLicense(licenseStream);
    }
}