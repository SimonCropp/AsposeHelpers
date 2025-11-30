using Aspose.Words.Loading;

namespace Aspose.Words;

public static partial class WordExtensions
{
    extension(Document)
    {
        /// <summary>
        /// Loads a MemoryStream into a <see cref="Document"/> and validates it is not corrupt.
        /// </summary>
        public static Document Load(MemoryStream stream, LoadFormat format = LoadFormat.Docx) =>
            Document.Load(
                stream,
                new LoadOptions
                {
                    LoadFormat = format
                });

        /// <summary>
        /// Loads a MemoryStream into a <see cref="Document"/> and validates it is not corrupt.
        /// </summary>
        public static Task<Document> Load(Stream stream, LoadFormat format = LoadFormat.Docx) =>
            Document.Load(
                stream,
                new LoadOptions
                {
                    LoadFormat = format
                });

        /// <summary>
        /// Loads a MemoryStream into a <see cref="Document"/> and validates it is not corrupt.
        /// </summary>
        public static async Task<Document> Load(Stream stream, LoadOptions options)
        {
            if (stream is MemoryStream memoryStream)
            {
                return Load(memoryStream, options);
            }

            memoryStream = new();
            await stream.CopyToAsync(memoryStream);
            return Load(memoryStream, options);
        }

        /// <summary>
        /// Loads a MemoryStream into a <see cref="Document"/> and validates it is not corrupt.
        /// </summary>
        public static Document Load(MemoryStream stream, LoadOptions options)
        {
            stream.Position = 0;
            var document = new Document(
                stream,
                options);

            stream.Position = 0;
            var fileFormat = FileFormatUtil.DetectFileFormat(stream);
            // DetectFileFormat uses the same logic as the Document but lets
            // us checked if it has mitakenly resolved to an incorrect format
            if (fileFormat.LoadFormat == options.LoadFormat)
            {
                return document;
            }

            throw new($"Bad document type or corrupt. Detected: {fileFormat.LoadFormat}. Expected: {options.LoadFormat}.");
        }
    }
}