using System.Drawing;
using Aspose.Cells;
using Aspose.Pdf.Devices;
using Aspose.Slides;
using Aspose.Words.Drawing;
using LoadOptions = Aspose.Slides.LoadOptions;
using Shape = Aspose.Words.Drawing.Shape;

namespace Aspose.Words;

public static class WordExtensions
{
    public static void WriteEmail(this DocumentBuilder builder, string email)
    {
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;
        builder.InsertHyperlink(email, $"mailto://{email}", false);
        builder.Font.ClearFormatting();
    }

    public static void WriteLink(this DocumentBuilder builder, string text, string link)
    {
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;
        builder.InsertHyperlink(text, link, false);
        builder.Font.ClearFormatting();
    }

    public static void WriteH1(this DocumentBuilder builder, string text) =>
        builder.WriteStyled(text, StyleIdentifier.Heading1);

    public static void WriteH2(this DocumentBuilder builder, string text) =>
        builder.WriteStyled(text, StyleIdentifier.Heading2);

    public static void WriteH3(this DocumentBuilder builder, string text) =>
        builder.WriteStyled(text, StyleIdentifier.Heading3);

    public static void WriteH4(this DocumentBuilder builder, string text) =>
        builder.WriteStyled(text, StyleIdentifier.Heading4);

    public static void WriteH5(this DocumentBuilder builder, string text) =>
        builder.WriteStyled(text, StyleIdentifier.Heading5);

    public static void SetMargins(this DocumentBuilder builder, double millimeters)
    {
        var margin = ConvertUtil.MillimeterToPoint(millimeters);
        var pageSetup = builder.PageSetup;
        pageSetup.TopMargin = margin;
        pageSetup.BottomMargin = margin;
        pageSetup.LeftMargin = margin;
        pageSetup.RightMargin = margin;
        pageSetup.HeaderDistance = margin;
        pageSetup.FooterDistance = margin;
    }

    public static void WriteStyled(this DocumentBuilder builder, string text, StyleIdentifier style)
    {
        builder.ParagraphFormat.StyleIdentifier = style;
        builder.Writeln(text);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
    }

    public static void ApplyBorder(this DocumentBuilder documentBuilder, LineStyle style)
    {
        var borders = documentBuilder.ParagraphFormat.Borders;
        borders[BorderType.Left].LineStyle = style;
        borders[BorderType.Right].LineStyle = style;
        borders[BorderType.Top].LineStyle = style;
        borders[BorderType.Bottom].LineStyle = style;
    }

    public static Shape InsertFullPageImage(this DocumentBuilder builder, Stream stream) =>
        builder.InsertImage(
            stream,
            RelativeHorizontalPosition.Margin,
            left: 0,
            RelativeVerticalPosition.Margin,
            top: 0,
            width: -1,
            height: -1,
            WrapType.Square);

    public static Shape InsertFullPageImage(this DocumentBuilder builder, string file) =>
        builder.InsertImage(
            file,
            RelativeHorizontalPosition.Margin,
            left: 0,
            RelativeVerticalPosition.Margin,
            top: 0,
            width: -1,
            height: -1,
            WrapType.Square);

    public static void AppendPresentation(this DocumentBuilder builder, Stream stream, Aspose.Slides.LoadFormat? format = null)
    {
        var options = new LoadOptions();
        if (format != null)
        {
            options.LoadFormat = format.Value;
        }

        using var presentation = new Presentation(stream, options);
        AppendPresentation(builder, presentation);
    }

    public static void AppendPresentation(this DocumentBuilder builder, string file, Aspose.Slides.LoadFormat? format = null)
    {
        var options = new LoadOptions();
        if (format != null)
        {
            options.LoadFormat = format.Value;
        }

        using var presentation = new Presentation(file, options);
        AppendPresentation(builder, presentation);
    }

    public static void AppendPresentation(this DocumentBuilder builder, Presentation presentation)
    {
        for (var index = 0; index < presentation.Slides.Count; index++)
        {
            using var htmlStream = new MemoryStream();
            presentation.Save(
                htmlStream,
                new[]
                {
                    index + 1
                },
                Slides.Export.SaveFormat.Html);
            htmlStream.Position = 0;
            builder.InsertHtml(Encoding.UTF8.GetString(htmlStream.GetBuffer()));
            if (index < presentation.Slides.Count - 1)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
        }
    }

    public static void AppendPdf(this DocumentBuilder builder, Stream stream)
    {
        using var document = new Aspose.Pdf.Document(stream);
        AppendPdf(builder, document);
    }

    public static void AppendPdf(this DocumentBuilder builder, string file)
    {
        using var document = new Aspose.Pdf.Document(file);
        AppendPdf(builder, document);
    }

    public static void AppendPdf(this DocumentBuilder builder, Pdf.Document document)
    {
        for (var index = 0; index < document.Pages.Count; index++)
        {
            var page = document.Pages[index + 1];
            using var imageStream = new MemoryStream();
            var pngDevice = new PngDevice();
            pngDevice.Process(page, imageStream);
            InsertFullPageImage(builder, imageStream);
            if (index < document.Pages.Count - 1)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
        }
    }

    public static void AppendWord(this DocumentBuilder builder, Stream stream, LoadFormat format = LoadFormat.Auto)
    {
        var options = new Loading.LoadOptions
        {
            LoadFormat = format
        };

        var document = new Document(stream, options);
        AppendWord(builder, document);
    }

    public static void AppendWord(this DocumentBuilder builder, string file, LoadFormat format = LoadFormat.Auto)
    {
        var options = new Loading.LoadOptions
        {
            LoadFormat = format
        };

        var document = new Document(file, options);
        AppendWord(builder, document);
    }

    public static void AppendWord(this DocumentBuilder builder, Document document)
    {
        var nestedBuilder = new DocumentBuilder(document);
        var setup = nestedBuilder.PageSetup;
        setup.PaperSize = builder.PageSetup.PaperSize;
        nestedBuilder.SetMargins(0);

        for (var index = 0; index < document.PageCount; index++)
        {
            var options = new Saving.ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new(index),
                OptimizeOutput = true
            };
            using var imageStream = new MemoryStream();
            document.Save(imageStream, options);
            InsertFullPageImage(builder, imageStream);
            if (index < document.PageCount - 1)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
        }
    }

    public static void AppendWorkbook(this DocumentBuilder builder, string file, Cells.LoadOptions? options = null)
    {
        using var book = new Workbook(file, options);
        AppendWorkbook(builder, book);
    }

    public static void AppendWorkbook(this DocumentBuilder builder, Stream stream, Cells.LoadOptions? options = null)
    {
        using var book = new Workbook(stream, options);
        AppendWorkbook(builder, book);
    }

    public static void AppendWorkbook(this DocumentBuilder builder, Workbook book)
    {
        var count = book.Worksheets.Count;
        for (var index = 0; index < count; index++)
        {
            var sheet = book.Worksheets[index];

            sheet.AutoFitColumns();
            sheet.Cells.DeleteBlankRows();
            sheet.Cells.DeleteBlankColumns();
            book.Worksheets.ActiveSheetName = sheet.Name;
            using var currentSheet = new MemoryStream();

            book.Save(
                currentSheet,
                new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = true,
                    ExportGridLines = true,
                });
            var html = Encoding.UTF8.GetString(currentSheet.GetBuffer());
            builder.InsertHtml(html);

            if (index < count - 1)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
        }
    }
}