namespace Aspose.Words;

public static partial class WordExtensions
{
    extension(Document document)
    {
        public void ModifyStyles(Action<Style> action)
        {
            foreach (var style in document.Styles)
            {
                action(style);
            }
        }

        public void ModifyStyleFonts(Action<Font> action)
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
}