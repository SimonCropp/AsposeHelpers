using Aspose.Words;

[TestFixture]
public class LoadDocxTests
{
    [Test]
    public void Load_ValidDocx_ReturnsDocument()
    {
        // Arrange
        using var stream = CreateValidDocxStream();

        // Act
        var document = WordExtensions.LoadDocx(stream);

        // Assert
        Assert.That(document, Is.Not.Null);
        // Stream position reset
        Assert.That(stream.Position, Is.EqualTo(0));
    }

    [Test]
    public void Load_ValidDocxWithContent_PreservesContent()
    {
        // Arrange
        using var stream = CreateValidDocxStream("Test content");

        // Act
        var document = WordExtensions.LoadDocx(stream);

        // Assert
        Assert.That(document.GetText().Trim(), Is.EqualTo("Test content"));
    }

    [Test]
    public void Load_DocFormat_ThrowsException()
    {
        // Arrange
        using var stream = CreateDocStream();

        // Act & Assert
        var exception = Assert.Throws<Exception>(() => WordExtensions.LoadDocx(stream))!;
        Assert.That(exception.Message, Does.Contain("Bad document type or corrupt"));
        Assert.That(exception.Message, Does.Contain("Expected: Docx"));
    }

    [Test]
    public void Load_RtfFormat_ThrowsException()
    {
        // Arrange
        using var stream = CreateRtfStream();

        // Act & Assert
        var exception = Assert.Throws<Exception>(() => WordExtensions.LoadDocx(stream))!;
        Assert.That(exception.Message, Does.Contain("Bad document type or corrupt"));
        Assert.That(exception.Message, Does.Contain("Detected: Rtf"));
    }

    [Test]
    public void Load_CorruptFile_ThrowsException()
    {
        // Arrange
        using var stream = new MemoryStream("Not a valid DOCX file"u8.ToArray());

        // Act & Assert
        var exception = Assert.Throws<Exception>(() => WordExtensions.LoadDocx(stream))!;
        Assert.That(exception.Message, Does.Contain("Bad document type or corrupt"));
    }

    [Test]
    public void Load_EmptyStream_ThrowsException()
    {
        // Arrange
        using var stream = new MemoryStream();

        // Act & Assert
        Assert.Throws<Exception>(() => WordExtensions.LoadDocx(stream));
    }

    [Test]
    public void Load_StreamNotAtBeginning_LoadsCorrectly()
    {
        // Arrange
        using var stream = CreateValidDocxStream();
        // Move position away from start
        stream.Position = 100;

        // Act
        var document = WordExtensions.LoadDocx(stream);

        // Assert
        Assert.That(document, Is.Not.Null);
        Assert.That(stream.Position, Is.EqualTo(0));
    }

    // Helper methods to create test streams
    static MemoryStream CreateValidDocxStream(string content = "")
    {
        var document = new Document();
        if (!string.IsNullOrEmpty(content))
        {
            var builder = new DocumentBuilder(document);
            builder.Write(content);
        }

        var stream = new MemoryStream();
        document.Save(stream, SaveFormat.Docx);
        stream.Position = 0;
        return stream;
    }

    static MemoryStream CreateDocStream()
    {
        var document = new Document();
        var stream = new MemoryStream();
        document.Save(stream, SaveFormat.Doc);
        stream.Position = 0;
        return stream;
    }

    static MemoryStream CreateRtfStream()
    {
        var document = new Document();
        var stream = new MemoryStream();
        document.Save(stream, SaveFormat.Rtf);
        stream.Position = 0;
        return stream;
    }
}