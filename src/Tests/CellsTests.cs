using Aspose.Cells;

[TestFixture]
public class CellsTests
{
    [Test]
    public async Task CellAlignment()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell = sheet.Cells["A1"];
        cell.PutValue("Hello World!");

        cell.AlignRight();
        cell.AlignTop();
        await Verify(cell.GetStyle());
    }

    [Test]
    public async Task SetColumnWidth()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell = sheet.Cells["A1"];
        cell.PutValue("Hello World!");
        cell.SetColumnWidth(10);
        await Verify(cell.GetStyle());
    }

    [Test]
    public async Task SheetAlignment()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell = sheet.Cells["A1"];
        cell.PutValue("Hello World!");

        sheet.AlignRight();
        sheet.AlignTop();
        await Verify(cell.GetStyle());
    }

    [Test]
    public async Task MakeHeadingsBoldSheet()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell = sheet.Cells["A1"];
        cell.PutValue("Hello World!");
        sheet.MakeHeadingsBold();
        await Verify(cell.GetStyle());
    }

    [Test]
    public async Task MakeHeadingsBoldBook()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell = sheet.Cells["A1"];
        cell.PutValue("Hello World!");
        workbook.MakeHeadingsBold();
        await Verify(cell.GetStyle());
    }

    [Test]
    public async Task FirstNullCell()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell = sheet.Cells["A1"];
        cell.PutValue("Hello World!");
        await Verify(sheet.FirstNullCell(0));
    }

    [Test]
    public async Task AddColumn()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AddColumn("Hello World!");
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDateTimeOffset()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(0, new DateTimeOffset(new(2020, 10, 7, 1, 2, 3)));
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDateTimeOffsetMinValue()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(0, DateTimeOffset.MinValue);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDate()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(0, new Date(2020, 10, 7));
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDateMinValue()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(0, Date.MinValue);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDateTime()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(0, new DateTime(2020, 10, 7, 1, 2, 3));
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDateTimeFormat()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(0, new DateTime(2020, 10, 7, 1, 2, 3), "yyyy");
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDateTimeMinValue()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(0, DateTime.MinValue);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellInt()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(0, 1);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellGuid()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(0, Guid.Empty);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellString()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(0, "The value");
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDecimal()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(0, (decimal)10);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellBool()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(0, true);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellHtml()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCellHtml(0, "<b>the value</b>");
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellHtml_NestedBug()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCellHtml(0, "<div> <div>AAA</div> </div>");
        await Verify(workbook);
    }
}