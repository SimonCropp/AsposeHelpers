using Aspose.Words;

[TestFixture]
public class WordTests
{
    [Test]
    public Task WriteEmail()
    {
        var builder = new DocumentBuilder();
        var pageSetup = builder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;

        #region WriteEmail

        builder.WriteEmail("the mail");

        #endregion

        return Verify(builder.Document);
    }

    [Test]
    public Task AddPageNumbers()
    {
        var builder = new DocumentBuilder();
        var pageSetup = builder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;

        builder.Writeln("the text");

        #region AddPageNumbers

        builder.AddPageNumbers();

        #endregion

        return Verify(builder.Document);
    }

    [Test]
    public Task InsertTocEntry()
    {
        var builder = new DocumentBuilder();

        builder.InsertTableOfContents("""TOC \o "2-3" \f \h \z \u""");

        builder.Writeln("the text");

        builder.InsertTocEntry("Custom toc entry", 2);

        builder.Document.UpdateFields();

        return Verify(builder.Document);
    }

    [Test]
    public Task WriteLink()
    {
        var documentBuilder = new DocumentBuilder();
        var pageSetup = documentBuilder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;

        #region WriteLink

        documentBuilder.WriteLink("the text", "the url");

        #endregion

        return Verify(documentBuilder.Document);
    }

    [Test]
    public Task TableAssignStyle()
    {
        var document = new Document();

        AddTableStyle(document);
        var builder = new DocumentBuilder(document);
        var table = builder.StartTable();
        builder.InsertCell();
        table.AssignStyle("MyStyle");
        builder.Write("Row 1, cell 1.");
        builder.EndRow();
        builder.EndTable();

        return Verify(document);
    }

    [Test]
    public Task WriteH1()
    {
        var documentBuilder = new DocumentBuilder();
        var pageSetup = documentBuilder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;

        #region WriteH1

        documentBuilder.WriteH1("the text");

        #endregion

        return Verify(documentBuilder.Document);
    }

    [Test]
    public Task WriteBoldItalic()
    {
        var builder = new DocumentBuilder();
        var pageSetup = builder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;

        builder.WriteBold("Bold");
        builder.WriteItalic("Italic");
        builder.WriteBoldItalic("BoldItalic");
        builder.Writeln();
        builder.WriteBoldLine("BoldLine");
        builder.WriteItalicLine("ItalicLine");
        builder.WriteBoldItalicLine("BoldItalicLine");

        return Verify(builder.Document);
    }

    [Test]
    public Task UseBoldItalic()
    {
        var builder = new DocumentBuilder();
        var pageSetup = builder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;

        using (builder.UseBold())
        {
            builder.Writeln("Bold");
        }

        using (builder.UseItalic())
        {
            builder.Writeln("Italic");
        }

        using (builder.UseBoldItalic())
        {
            builder.Writeln("BoldItalic");
        }

        return Verify(builder.Document);
    }

    [Test]
    public Task WriteNamedStyle()
    {
        var document = new Document();

        AddParagraphStyle(document);
        var documentBuilder = new DocumentBuilder(document);
        var pageSetup = documentBuilder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;

        documentBuilder.WriteStyled("the text", "MyStyle");

        return Verify(documentBuilder.Document);
    }

    [Test]
    public Task UseStyled()
    {
        var document = new Document();

        AddParagraphStyle(document);
        var documentBuilder = new DocumentBuilder(document);
        var pageSetup = documentBuilder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;

        documentBuilder.Writeln("one");
        using (documentBuilder.UseStyled("MyStyle"))
        {
            documentBuilder.Writeln("two");
        }

        documentBuilder.Writeln("three");
        return Verify(documentBuilder.Document);
    }

    static void AddTableStyle(Document document)
    {
        var style = document.Styles.Add(StyleType.Table, "MyStyle");
        style.Font.Size = 24;
        style.Font.Name = "Verdana";
        style.ParagraphFormat.SpaceAfter = 12;
    }

    static void AddParagraphStyle(Document document)
    {
        var style = document.Styles.Add(StyleType.Paragraph, "MyStyle");
        style.Font.Size = 24;
        style.Font.Name = "Verdana";
        style.ParagraphFormat.SpaceAfter = 12;
    }

    [Test]
    public Task ModifyStyles()
    {
        var documentBuilder = new DocumentBuilder();
        var pageSetup = documentBuilder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;
        var document = documentBuilder.Document;

        documentBuilder.WriteH2("the text");

        #region ModifyStyles

        document.ModifyStyles(_ =>
            _.Font?.Italic = false);

        #endregion

        return Verify(document);
    }

    [Test]
    public Task ModifyStyleFonts()
    {
        var documentBuilder = new DocumentBuilder();
        var pageSetup = documentBuilder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;
        var document = documentBuilder.Document;

        documentBuilder.WriteH2("the text");

        #region ModifyStyleFonts

        document.ModifyStyleFonts(_ =>
        {
            _.Italic = false;
        });

        #endregion

        return Verify(document);
    }

    [Test]
    public Task SetMargins()
    {
        var documentBuilder = new DocumentBuilder();
        var pageSetup = documentBuilder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;

        documentBuilder.WriteH1("the text");

        #region SetMargins

        documentBuilder.SetMargins(10);

        #endregion

        return Verify(documentBuilder.Document);
    }

    [Test]
    public Task ApplyBorder()
    {
        var documentBuilder = new DocumentBuilder();
        var pageSetup = documentBuilder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;

        #region ApplyBorder

        documentBuilder.ApplyBorder(LineStyle.Thick);
        documentBuilder.Write("some text");

        #endregion

        return Verify(documentBuilder.Document);
    }
}