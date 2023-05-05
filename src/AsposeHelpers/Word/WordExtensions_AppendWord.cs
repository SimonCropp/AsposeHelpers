using Aspose.Words.Drawing;

namespace Aspose.Words;

public static partial class WordExtensions
{
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
        for (var index = 0; index < document.PageCount; index++)
        {
            using var imageStream = ToImage(document, index);
            var image = builder.InsertImage(imageStream);
            image.WrapType = WrapType.Square;
            if (index < document.PageCount - 1)
            {
                builder.InsertBreak(BreakType.SectionBreakNewPage);
            }
        }
    }

    public static MemoryStream ToImage(this Document document, int page)
    {
        var options = new Saving.ImageSaveOptions(SaveFormat.Png)
        {
            PageSet = new(page),
            OptimizeOutput = true
        };
        var stream = new MemoryStream();
        document.Save(stream, options);
        return stream;
    }
}