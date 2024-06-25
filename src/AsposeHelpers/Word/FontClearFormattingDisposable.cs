namespace Aspose.Words;

class FontClearFormattingDisposable(DocumentBuilder builder) :
    IDisposable
{
    public void Dispose() =>
        builder.Font.ClearFormatting();
}