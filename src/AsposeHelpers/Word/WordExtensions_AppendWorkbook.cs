using Aspose.Cells;
using Aspose.Words.Drawing;
using ImageType = Aspose.Cells.Drawing.ImageType;

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
            using var imageStream = new MemoryStream();

            book.Save(
                imageStream,
                new ImageSaveOptions
                {
                    ImageOrPrintOptions =
                    {
                        ImageType = ImageType.Png
                    }
                });
            var image = builder.InsertImage(imageStream);
            image.WrapType = WrapType.Square;

            if (index < count - 1)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
        }
    }
}