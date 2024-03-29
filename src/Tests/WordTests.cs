﻿using Aspose.Words;

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
    public Task ModifyStyles()
    {
        var documentBuilder = new DocumentBuilder();
        var pageSetup = documentBuilder.PageSetup;
        pageSetup.PaperSize = PaperSize.A5;
        var document = documentBuilder.Document;

        documentBuilder.WriteH2("the text");

        #region ModifyStyles

        document.ModifyStyles(_ =>
        {
            if (_.Font != null)
            {
                _.Font.Italic = false;
            }
        });

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