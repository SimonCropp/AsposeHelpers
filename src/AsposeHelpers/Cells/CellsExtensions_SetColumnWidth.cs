namespace Aspose.Cells;

public static partial class CellsExtensions
{
    public static void SetColumnWidth(this Worksheet sheet, int column, double width) =>
        sheet.Cells.SetColumnWidth(column, width);

    public static void SetColumnWidth(this Cell cell, double width) =>
        cell.Worksheet.Cells.SetColumnWidth(cell.Column, width);
}