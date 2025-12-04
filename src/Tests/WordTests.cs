using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Fields;

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
        var builder = new DocumentBuilder(document);
        var pageSetup = builder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;

        builder.WriteStyled("the text", "MyStyle");

        return Verify(builder.Document);
    }

    [Test]
    public Task UseStyled()
    {
        var document = new Document();

        AddParagraphStyle(document);
        var builder = new DocumentBuilder(document);
        var pageSetup = builder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;

        builder.Writeln("one");
        using (builder.UseStyled("MyStyle"))
        {
            builder.Writeln("two");
        }

        builder.Writeln("three");
        return Verify(builder.Document);
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
        var builder = new DocumentBuilder();
        var pageSetup = builder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;
        var document = builder.Document;

        builder.WriteH2("the text");

        #region ModifyStyles

        document.ModifyStyles(_ =>
            _.Font?.Italic = false);

        #endregion

        return Verify(document);
    }

    [Test]
    public Task ModifyStyleFonts()
    {
        var builder = new DocumentBuilder();
        var pageSetup = builder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;
        var document = builder.Document;

        builder.WriteH2("the text");

        #region ModifyStyleFonts

        document.ModifyStyleFonts(_ =>
            _.Italic = false);

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

    [Test]
    public void ReplaceField_ReplacesFieldWithValue()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("TestField", TextFormFieldType.Regular, "", "", 0);

        builder.ReplaceField("TestField", "Replacement Text");

        var field = document.Range.FormFields.SingleOrDefault(_ => _.Name == "TestField");
        Assert.That(field, Is.Null);
        Assert.That(document.Range.Text, Does.Contain("Replacement Text"));
    }

    [Test]
    public void ReplaceField_ReplacesMultipleFieldsWithSameName()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("TestField", TextFormFieldType.Regular, "", "", 0);
        builder.Write(" some text ");
        builder.InsertTextInput("TestField", TextFormFieldType.Regular, "", "", 0);
        builder.Write(" more text ");
        builder.InsertTextInput("TestField", TextFormFieldType.Regular, "", "", 0);

        builder.ReplaceField("TestField", "Replacement Text");

        var fields = document.Range.FormFields.Where(_ => _.Name == "TestField").ToList();
        Assert.That(fields, Is.Empty);

        var text = document.Range.Text;
        var count = Regex.Matches(text, "Replacement Text").Count;
        Assert.That(count, Is.EqualTo(3));
    }

    [Test]
    public void ReplaceField_ThrowsWhenFieldNotFound()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        var exception = Assert.Throws<Exception>(() =>
            builder.ReplaceField("NonExistent", "value"))!;

        Assert.That(exception.Message, Is.EqualTo("Could not find field: NonExistent"));
    }

    [Test]
    public void DisplaceField_RemovesField()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("TestField", TextFormFieldType.Regular, "", "", 0);
        Assert.That(document.Range.FormFields.SingleOrDefault(_ => _.Name == "TestField"), Is.Not.Null);

        builder.DisplaceField("TestField");

        var field = document.Range.FormFields.SingleOrDefault(_ => _.Name == "TestField");
        Assert.That(field, Is.Null);
    }

    [Test]
    public void DisplaceField_ThrowsWhenFieldNotFound()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        var exception = Assert.Throws<Exception>(() =>
            builder.DisplaceField("NonExistent"))!;

        Assert.That(exception.Message, Is.EqualTo("Could not find field: NonExistent"));
    }

    [Test]
    public void ReplaceField_HandlesMultipleFields()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("Field1", TextFormFieldType.Regular, "", "", 0);
        builder.Write(" ");
        builder.InsertTextInput("Field2", TextFormFieldType.Regular, "", "", 0);

        builder.ReplaceField("Field1", "First");
        builder.ReplaceField("Field2", "Second");

        var fields = document.Range.FormFields;
        Assert.That(fields.SingleOrDefault(_ => _.Name == "Field1"), Is.Null);
        Assert.That(fields.SingleOrDefault(_ => _.Name == "Field2"), Is.Null);
        var text = document.Range.Text;
        Assert.That(text, Does.Contain("First"));
        Assert.That(text, Does.Contain("Second"));
    }
}