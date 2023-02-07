using Aspose.Slides;
using LoadOptions = Aspose.Slides.LoadOptions;

namespace Aspose.Words;

public static partial class WordExtensions
{
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
}