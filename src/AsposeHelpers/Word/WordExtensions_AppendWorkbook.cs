using Aspose.Cells;

namespace Aspose.Words;

public static partial class WordExtensions
{
    public static void AppendWorkbook(this DocumentBuilder builder, string file, LoadOptions? options = null)
    {
        using var book = new Workbook(file, options);
        AppendWorkbook(builder, book);
    }

    public static void AppendWorkbook(this DocumentBuilder builder, Stream stream, LoadOptions? options = null)
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