namespace Aspose.Words;

public static partial class WordExtensions
{
    extension(Document)
    {
        /// <summary>
        /// Loads a MemoryStream into a <see cref="Document"/> and validates it is not corrupt.
        /// </summary>
        public static Document LoadDocx(MemoryStream stream)
        {
            stream.Position = 0;

            var document = new Document(
                stream,
                new()
                {
                    LoadFormat = LoadFormat.Docx
                });

            stream.Position = 0;
            var fileFormat = FileFormatUtil.DetectFileFormat(stream);
            // DetectFileFormat uses the same logic as the Document but lets
            // us checked if it has mitakenly resolved to an incorrect format
            var format = fileFormat.LoadFormat;
            if (format != LoadFormat.Docx)
            {
                throw new($"Bad document type or corrupt. Detected: {format}. Expected: Docx.");
            }

            return document;
        }
    }
}