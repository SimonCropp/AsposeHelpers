using Aspose.Cells;

[TestFixture]
public class CellsExtensionsTests
{
    [Test]
    public void ReadXlsx_ValidXlsx_ReturnsWorkbook()
    {
        // Arrange
        using var stream = CreateValidXlsxStream();

        // Act
        var workbook = CellsExtensions.ReadXlsx(stream);

        // Assert
        Assert.That(workbook, Is.Not.Null);
        Assert.That(stream.Position, Is.EqualTo(0)); // Stream position reset
    }

    [Test]
    public void ReadXlsx_ValidXlsxWithData_PreservesContent()
    {
        // Arrange
        using var stream = CreateValidXlsxStream("Test Data", "Sheet1");

        // Act
        var workbook = CellsExtensions.ReadXlsx(stream);

        // Assert
        var worksheet = workbook.Worksheets[0];
        Assert.That(worksheet.Cells["A1"].StringValue, Is.EqualTo("Test Data"));
        Assert.That(worksheet.Name, Is.EqualTo("Sheet1"));
    }

    [Test]
    public void ReadXlsx_ValidXlsxWithMultipleSheets_PreservesAllSheets()
    {
        // Arrange
        using var stream = CreateXlsxWithMultipleSheets();

        // Act
        var workbook = CellsExtensions.ReadXlsx(stream);

        // Assert
        Assert.That(workbook.Worksheets.Count, Is.EqualTo(3));
        Assert.That(workbook.Worksheets[0].Name, Is.EqualTo("Sheet1"));
        Assert.That(workbook.Worksheets[1].Name, Is.EqualTo("Sheet2"));
        Assert.That(workbook.Worksheets[2].Name, Is.EqualTo("Sheet3"));
    }

    [Test]
    public void ReadXlsx_XlsFormat_ThrowsException()
    {
        // Arrange
        using var stream = CreateXlsStream();

        // Act & Assert
        var exception = Assert.Throws<Exception>(() => CellsExtensions.ReadXlsx(stream))!;
        Assert.That(exception.Message, Does.Contain("Bad document type or corrupt"));
        Assert.That(exception.Message, Does.Contain("Expected: Xlsx"));
    }

    [Test]
    public void ReadXlsx_CsvFormat_ThrowsException()
    {
        // Arrange
        using var stream = CreateCsvStream();

        // Act & Assert
        var exception = Assert.Throws<CellsException>(() => CellsExtensions.ReadXlsx(stream))!;
        Assert.That(exception.Message, Does.Contain("Unsupported file format"));
    }

    [Test]
    public void ReadXlsx_CorruptFile_ThrowsException()
    {
        // Arrange
        using var stream = new MemoryStream("Not a valid XLSX file"u8.ToArray());

        // Act & Assert
        var exception = Assert.Throws<CellsException>(() => CellsExtensions.ReadXlsx(stream))!;
        Assert.That(exception.Message, Does.Contain("File is corrupted"));
    }

    [Test]
    public void ReadXlsx_EmptyStream_ThrowsException()
    {
        // Arrange
        using var stream = new MemoryStream();

        // Act & Assert
        Assert.Throws<Exception>(() => CellsExtensions.ReadXlsx(stream));
    }

    [Test]
    public void ReadXlsx_StreamNotAtBeginning_LoadsCorrectly()
    {
        // Arrange
        using var stream = CreateValidXlsxStream("Test");
        // Move position away from start
        stream.Position = 100;

        // Act
        var workbook = CellsExtensions.ReadXlsx(stream);

        // Assert
        Assert.That(workbook, Is.Not.Null);
        Assert.That(stream.Position, Is.EqualTo(0));
        Assert.That(workbook.Worksheets[0].Cells["A1"].StringValue, Is.EqualTo("Test"));
    }

    // Helper methods to create test streams
    static MemoryStream CreateValidXlsxStream(string? cellValue = null, string? sheetName = null)
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        if (sheetName != null)
        {
            worksheet.Name = sheetName;
        }

        if (cellValue != null)
        {
            worksheet.Cells["A1"].PutValue(cellValue);
        }

        var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Xlsx);
        stream.Position = 0;
        return stream;
    }

    static MemoryStream CreateXlsxWithMultipleSheets()
    {
        var workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Xlsx);
        stream.Position = 0;
        return stream;
    }

    static MemoryStream CreateXlsStream()
    {
        var workbook = new Workbook();
        var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Excel97To2003);
        stream.Position = 0;
        return stream;
    }

    static MemoryStream CreateCsvStream()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Test");
        var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Csv);
        stream.Position = 0;
        return stream;
    }
}