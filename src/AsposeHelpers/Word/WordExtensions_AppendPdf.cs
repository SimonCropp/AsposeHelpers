﻿using Aspose.Pdf.Devices;
using Aspose.Words.Drawing;

namespace Aspose.Words;

public static partial class WordExtensions
{
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
            using var imageStream = ToImage(document, index);
            var image = builder.InsertImage(imageStream);
            image.WrapType = WrapType.Square;
            if (index < document.Pages.Count - 1)
            {
                builder.InsertBreak(BreakType.SectionBreakNewPage);
            }
        }
    }

    public static MemoryStream ToImage(this Pdf.Document document, int index)
    {
        var page = document.Pages[index + 1];
        var stream = new MemoryStream();
        var pngDevice = new PngDevice();
        pngDevice.Process(page, stream);
        return stream;
    }
}