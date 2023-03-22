using Aspose.Words;
using Document = Aspose.Words.Document;
using Task = System.Threading.Tasks.Task;

[TestFixture]
public class WordTests
{
    [Test]
    public Task AppendDocument()
    {
        var lic = new License();
        lic.SetLicense(@"C:\Code\MinisterialProgramManager\src\Manager\Aspose.Total.lic");
        var document = new Document();

        var builder = new DocumentBuilder(document);
        var setup = builder.PageSetup;
        setup.PaperSize = PaperSize.A4;
        setup.Margins = Margins.Narrow;
        setup.Orientation = Orientation.Portrait;
        var path = @"C:\Code\inputdocs\word.docx";
        foreach (var process in Process.GetProcessesByName("WINWORD"))
        {
            process.Kill();
            process.WaitForExit();
        }
        File.Delete(path);
        File.Delete(@"C:\Code\inputdocs\~$word.docx");
        foreach (var file in Directory.EnumerateFiles(@"C:\Code\inputdocs"))
        {
            setup.ClearFormatting();
            builder.WriteH1(file);
            var input = new Document(file);
            var nestedBuilder = new DocumentBuilder(document);
            var nestedSetup = nestedBuilder.PageSetup;
            nestedSetup.PaperSize = setup.PaperSize;
            nestedSetup.Margins = setup.Margins;
            nestedSetup.Orientation = setup.Orientation;

            builder.InsertDocument(input, ImportFormatMode.UseDestinationStyles);
            builder.InsertBreak(BreakType.PageBreak);
        }

        document.Save(path);
        //return Verify(document);
        var startInfo = new ProcessStartInfo(path)
        {
            UseShellExecute = true
        };
        Process.Start(startInfo);
        return Task.CompletedTask;
    }

}