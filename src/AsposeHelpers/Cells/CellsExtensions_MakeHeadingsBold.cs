namespace Aspose.Cells;

public static partial class CellsExtensions
{
    public static void MakeHeadingsBold(this Worksheet sheet)
    {
        var firstRow = sheet.Cells.Rows[0];
        foreach (Cell cell in firstRow)
        {
            var style = cell.GetStyle();
            style.Font.IsBold = true;
            cell.SetStyle(style);
        }
    }
    public static void MakeHeadingsBold(this Workbook book)
    {
        foreach (var sheet in book.Worksheets)
        {
            sheet.MakeHeadingsBold();
        }
    }
}