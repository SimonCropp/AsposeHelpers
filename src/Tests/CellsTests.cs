using Aspose.Cells;

[TestFixture]
public class CellsTests
{
    [Test]
    public async Task MakeHeadingsBoldSheet()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell = sheet.Cells["A1"];
        cell.PutValue("Hello World!");
        sheet.MakeHeadingsBold();
        await Verify(workbook);
    }
    [Test]
    public async Task MakeHeadingsBoldBook()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell = sheet.Cells["A1"];
        cell.PutValue("Hello World!");
        workbook.MakeHeadingsBold();
        await Verify(workbook);
    }
}