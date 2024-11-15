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
        sheet.AddColumn("Hello World1", 100);
        sheet.AddColumn("Hello World2", 100);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDateTimeOffset()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell = sheet.AppendCell(1, new DateTimeOffset(new(2020, 10, 7, 1, 2, 3)));
        cell.SetColumnWidth(50);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDateTimeOffsetMinValue()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell = sheet.AppendCell(1, DateTimeOffset.MinValue);
        cell.SetColumnWidth(50);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDate()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell =   sheet.AppendCell(1, new Date(2020, 10, 7));
        cell.SetColumnWidth(50);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDateMinValue()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell =   sheet.AppendCell(1, Date.MinValue);
        cell.SetColumnWidth(50);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDateTime()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell =   sheet.AppendCell(1, new DateTime(2020, 10, 7, 1, 2, 3));
        cell.SetColumnWidth(50);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDateTimeFormat()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell = sheet.AppendCell(1, new DateTime(2020, 10, 7, 1, 2, 3), "yyyy");
        cell.SetColumnWidth(50);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellDateTimeMinValue()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(1, DateTime.MinValue);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellInt()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(1, 1);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellGuid()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        sheet.AppendCell(1, Guid.Empty);
        await Verify(workbook);
    }

    [Test]
    public async Task AppendCellString()
    {
        using var book = new Workbook();

        var sheet = book.Worksheets[0];
        sheet.AppendCell(1, "The value");
        await Verify(book);
    }

    [Test]
    public async Task AppendLinkCell()
    {
        using var book = new Workbook();

        var sheet = book.Worksheets[0];
        sheet.AppendLinkCell(1, "https://google.com", "The value");
        await Verify(book);
    }

    [Test]
    public async Task AppendLinkCellWithText()
    {
        using var book = new Workbook();

        var sheet = book.Worksheets[0];
        sheet.AppendLinkCell(1, "https://google.com", "The value");
        await Verify(book);
    }

    [Test]
    public async Task AppendCellDecimal()
    {
        using var book = new Workbook();

        var sheet = book.Worksheets[0];
        sheet.AppendCell(1, (decimal)10);
        await Verify(book);
    }

    [Test]
    public async Task AppendCellBool()
    {
        using var book = new Workbook();

        var sheet = book.Worksheets[0];
        sheet.AppendCell(1, true);
        await Verify(book);
    }

    [Test]
    public async Task AppendCellHtml()
    {
        using var book = new Workbook();

        var sheet = book.Worksheets[0];
        sheet.AppendCellHtml(1, "<b>the value</b>");
        await Verify(book);
    }

    [Test]
    public async Task AppendCellHtml_NestedBug()
    {
        using var book = new Workbook();

        var sheet = book.Worksheets[0];
        sheet.AppendCellHtml(1, "<div> <div>AAA</div> </div>");
        await Verify(book);
    }
}