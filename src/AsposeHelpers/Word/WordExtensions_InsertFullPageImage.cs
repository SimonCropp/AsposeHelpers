using Aspose.Words.Drawing;
using Shape = Aspose.Words.Drawing.Shape;

namespace Aspose.Words;

public static partial class WordExtensions
{
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
}