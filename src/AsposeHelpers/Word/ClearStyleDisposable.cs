namespace Aspose.Words;

class ClearStyleDisposable(DocumentBuilder builder) :
    IDisposable
{
    public void Dispose()
    {
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Font.ClearFormatting();
    }
}