namespace Aspose.Words;

public static partial class WordExtensions
{
    public static void ModifyStyles(this Document document, Action<Style> action)
    {
        foreach (var style in document.Styles)
        {
            action(style);
        }
    }

    public static void ModifyStyleFonts(this Document document, Action<Font> action)
    {
        foreach (var style in document.Styles)
        {
            if (style.Font != null)
            {
                action(style.Font);
            }
        }
    }
}