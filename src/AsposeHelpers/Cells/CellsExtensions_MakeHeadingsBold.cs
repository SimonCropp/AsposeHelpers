namespace Aspose.Cells;

public static partial class CellsExtensions
{
    extension(Worksheet sheet)
    {
        public void MakeHeadingsBold()
        {
            var firstRow = sheet.Cells.Rows[0];
            foreach (Cell cell in firstRow)
            {
                var style = cell.GetStyle();
                style.Font.IsBold = true;
                cell.SetStyle(style);
            }
        }
    }
    extension(Workbook book)
    {
        public void MakeHeadingsBold()
        {
            foreach (var sheet in book.Worksheets)
            {
                sheet.MakeHeadingsBold();
            }
        }
    }
}