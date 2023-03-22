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
        document.RemoveAllChildren();

        var builder = new DocumentBuilder(document);
        var path = @"C:\Code\inputdocs\word.docx";
        foreach (var process in Process.GetProcessesByName("WINWORD"))
        {
            process.Kill();
        }
        document.Save(path);
        File.Delete(path);
        foreach (var file in Directory.EnumerateFiles(@"C:\Code\inputdocs"))
        {
            builder.PageSetup.ClearFormatting();
            builder.WriteH1(file);
            var input = new Document(file);
            builder.CurrentSection.PageSetup.Orientation = Orientation.Portrait;
            builder.InsertDocument(input, ImportFormatMode.KeepSourceFormatting);
            builder.PageSetup.ClearFormatting();
            builder.InsertBreak(BreakType.PageBreak);
            // builder.PageSetup.Orientation = Orientation.Portrait;
            // builder.PageSetup.PaperSize = PaperSize.A4;
        }

        //return Verify(document);
        var startInfo = new ProcessStartInfo(path)
        {
            UseShellExecute = true
        };
        Process.Start(startInfo);
        return Task.CompletedTask;
    }

}