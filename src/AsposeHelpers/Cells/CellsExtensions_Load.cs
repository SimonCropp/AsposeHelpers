namespace Aspose.Cells;

public static partial class CellsExtensions
{
    extension(Cell)
    {
        /// <summary>
        /// Loads a MemoryStream into a <see cref="Workbook"/> and validates it is not corrupt.
        /// </summary>
        public static Workbook ReadXlsx(MemoryStream stream)
        {
            stream.Position = 0;
            var book = new Workbook(stream, new(LoadFormat.Xlsx));

            stream.Position = 0;
            var fileFormat = FileFormatUtil.DetectFileFormat(stream);
            // DetectFileFormat uses the same logic as the Document but lets
            // us checked if it has mitakenly resolved to an incorrect format
            var format = fileFormat.LoadFormat;
            if (format != LoadFormat.Xlsx)
            {
                throw new($"Bad document type or corrupt. Detected: {format}. Expected: Xlsx.");
            }

            return book;
        }
    }
}