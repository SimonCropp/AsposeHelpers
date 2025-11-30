namespace Aspose.Cells;

public static partial class CellsExtensions
{
    extension(Worksheet sheet)
    {
        public void SetColumnWidth(int column, double width) =>
            sheet.Cells.SetColumnWidth(column, width);
    }

    extension(Cell cell)
    {
        public void SetColumnWidth(double width) =>
            cell.Worksheet.Cells.SetColumnWidth(cell.Column, width);
    }
}