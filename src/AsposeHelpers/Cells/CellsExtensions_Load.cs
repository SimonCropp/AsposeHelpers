namespace Aspose.Cells;

public static partial class CellsExtensions
{
    extension(Cell)
    {
        /// <summary>
        /// Loads a MemoryStream into a <see cref="Workbook"/> and validates it is not corrupt.
        /// </summary>
        public static Workbook Load(MemoryStream stream, LoadFormat format = LoadFormat.Xlsx) =>
            Cell.Load(stream, new LoadOptions(format));

        /// <summary>
        /// Loads a MemoryStream into a <see cref="Workbook"/> and validates it is not corrupt.
        /// </summary>
        public static Task<Workbook> Load(Stream stream, LoadFormat format = LoadFormat.Xlsx) =>
            Cell.Load(stream, new LoadOptions(format));

        /// <summary>
        /// Loads a MemoryStream into a <see cref="Workbook"/> and validates it is not corrupt.
        /// </summary>
        public static async Task<Workbook> Load(Stream stream, LoadOptions options)
        {
            if (stream is MemoryStream memoryStream)
            {
                return Load(memoryStream, options);
            }

            var destination = new MemoryStream();
            await stream.CopyToAsync(destination);
            return Load(destination, options);
        }

        /// <summary>
        /// Loads a MemoryStream into a <see cref="Workbook"/> and validates it is not corrupt.
        /// </summary>
        public static Workbook Load(MemoryStream stream, LoadOptions options)
        {
            stream.Position = 0;
            var book = new Workbook(stream, options);

            stream.Position = 0;
            var fileFormat = FileFormatUtil.DetectFileFormat(stream);
            // DetectFileFormat uses the same logic as the Document but lets
            // us checked if it has mitakenly resolved to an incorrect format
            if (fileFormat.LoadFormat == options.LoadFormat)
            {
                return book;
            }

            throw new($"Bad document type or corrupt. Detected: {fileFormat.LoadFormat}. Expected: {options.LoadFormat}.");
        }
    }
}