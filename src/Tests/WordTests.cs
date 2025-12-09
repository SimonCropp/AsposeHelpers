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
    //TODO: should not be possible to have multiple form fields with the same name
    [Explicit]
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
    public Task ReplaceField_ThrowsWhenFieldNotFound()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        var exception = Assert.Throws<Exception>(() =>
            builder.ReplaceField("NonExistent", "value"))!;

        return Verify(exception)
            .IgnoreStackTrace();
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
    public Task DisplaceField_ThrowsWhenFieldNotFound()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        var exception = Assert.Throws<Exception>(() =>
            builder.DisplaceField("NonExistent"))!;

        return Verify(exception)
            .IgnoreStackTrace();
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

    [Test]
    public void FindField_WhenFieldExists_ReturnsField()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("TestField", TextFormFieldType.Regular, "", "", 0);

        var result = builder.FindField("TestField");

        Assert.That(result.Name, Is.EqualTo("TestField"));
    }

    [Test]
    public void TryFindField_WhenFieldExists_ReturnsTrue()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("TestField", TextFormFieldType.Regular, "", "", 0);

        var found = builder.TryFindField("TestField", out var result);

        Assert.That(found, Is.True);
        Assert.That(result, Is.Not.Null);
        Assert.That(result!.Name, Is.EqualTo("TestField"));
    }

    [Test]
    public void TryFindField_WhenFieldNotFound_ReturnsFalse()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("TestField", TextFormFieldType.Regular, "", "", 0);

        var found = builder.TryFindField("NonExistent", out var result);

        Assert.That(found, Is.False);
        Assert.That(result, Is.Null);
    }

    [Test]
    public void TryFindField_WhenDocumentHasNoFields_ReturnsFalse()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        var found = builder.TryFindField("TestField", out var result);

        Assert.That(found, Is.False);
        Assert.That(result, Is.Null);
    }

    [Test]
    public Task FindField_WhenDocumentHasNoFields_Throws()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        var exception = Assert.Throws<Exception>(() =>
            builder.FindField("TestField"))!;

        return Verify(exception)
            .IgnoreStackTrace();
    }

    [Test]
    public Task FindField_WhenFieldNotFound_ThrowsWithFieldList()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("FirstName", TextFormFieldType.Regular, "", "", 0);
        builder.InsertTextInput("LastName", TextFormFieldType.Regular, "", "", 0);
        builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "", 0);

        var exception = Assert.Throws<Exception>(() =>
            builder.FindField("MissingField"))!;

        return Verify(exception)
            .IgnoreStackTrace();
    }

    [Test]
    public void FindField_WithExactMatch_DoesNotReturnPartialMatches()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 0);
        builder.InsertTextInput("FullName", TextFormFieldType.Regular, "", "", 0);
        builder.InsertTextInput("FirstName", TextFormFieldType.Regular, "", "", 0);

        var result = builder.FindField("Name");

        Assert.That(result.Name, Is.EqualTo("Name"));
    }

    [Test]
    public Task FindField_WhenFieldNotFound_ListsDistinctFieldsOnly()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "", 0);
        builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "", 0);
        builder.InsertTextInput("Phone", TextFormFieldType.Regular, "", "", 0);
        builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "", 0);

        var exception = Assert.Throws<Exception>(() =>
            builder.FindField("MissingField"))!;

        return Verify(exception)
            .IgnoreStackTrace();
    }

    [Test]
    public Task FindField_WhenFieldNotFound_ListsFieldsAlphabetically()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("Zebra", TextFormFieldType.Regular, "", "", 0);
        builder.InsertTextInput("Apple", TextFormFieldType.Regular, "", "", 0);
        builder.InsertTextInput("Mango", TextFormFieldType.Regular, "", "", 0);
        builder.InsertTextInput("Banana", TextFormFieldType.Regular, "", "", 0);

        var exception = Assert.Throws<Exception>(() =>
            builder.FindField("MissingField"))!;

        return Verify(exception)
            .IgnoreStackTrace();
    }

    [Test]
    public void FindFieldByValue_MatchesByResult_ReturnsSingleField()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("TestField", TextFormFieldType.Regular, "", "Expected Value", 0);

        var result = builder.FindFieldByValue("Expected Value");

        Assert.That(result, Has.Count.EqualTo(1));
        Assert.That(result[0].Result, Is.EqualTo("Expected Value"));
    }

    [Test]
    public void FindFieldByValue_MatchesByResult_ReturnsMultipleFieldsWithSameResult()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("Field1", TextFormFieldType.Regular, "", "Same Value", 0);
        builder.Write(" ");
        builder.InsertTextInput("Field2", TextFormFieldType.Regular, "", "Same Value", 0);
        builder.Write(" ");
        builder.InsertTextInput("Field3", TextFormFieldType.Regular, "", "Different Value", 0);

        var result = builder.FindFieldByValue("Same Value");

        Assert.That(result, Has.Count.EqualTo(2));
        Assert.That(result.All(_ => _.Result == "Same Value"), Is.True);
    }

    [Test]
    public void ReplaceField_MatchesByResult_ReplacesField()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("TestField", TextFormFieldType.Regular, "", "Original Value", 0);

        builder.ReplaceField("Original Value", "Replacement Text");

        var field = document.Range.FormFields.SingleOrDefault(_ => _.Name == "TestField");
        Assert.That(field, Is.Null);
        Assert.That(document.Range.Text, Does.Contain("Replacement Text"));
    }

    [Test]
    public void ReplaceField_MatchesByResult_ReplacesMultipleFields()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.InsertTextInput("Field1", TextFormFieldType.Regular, "", "Shared Value", 0);
        builder.Write(" some text ");
        builder.InsertTextInput("Field2", TextFormFieldType.Regular, "", "Shared Value", 0);
        builder.Write(" more text ");
        builder.InsertTextInput("Field3", TextFormFieldType.Regular, "", "Shared Value", 0);

        builder.ReplaceField("Shared Value", "New Text");

        var fields = document.Range.FormFields;
        Assert.That(fields, Is.Empty);

        var text = document.Range.Text;
        var count = Regex.Matches(text, "New Text").Count;
        Assert.That(count, Is.EqualTo(3));
    }

    [Test]
    public void ReplaceField_PrefersNameOverValue()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        // Field with name "SearchTerm"
        builder.InsertTextInput("SearchTerm", TextFormFieldType.Regular, "", "Other Value", 0);
        builder.Write(" ");
        // Field with result "SearchTerm"
        builder.InsertTextInput("Field2", TextFormFieldType.Regular, "", "SearchTerm", 0);

        builder.ReplaceField("SearchTerm", "Replaced");

        // Should only replace the field named "SearchTerm", not the one with result "SearchTerm"
        var field2 = document.Range.FormFields.SingleOrDefault(_ => _.Name == "Field2")!;
        Assert.That(field2.Result, Is.EqualTo("SearchTerm"));
        Assert.That(document.Range.Text, Does.Contain("Replaced"));
    }
}