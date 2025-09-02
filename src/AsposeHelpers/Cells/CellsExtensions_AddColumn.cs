namespace Aspose.Cells;

public static partial class CellsExtensions
{
    public static int AddColumn(this Worksheet sheet, string heading, int? width = null, Action<Style>? modifyStyle = null)
    {
        var cells = sheet.Cells;
        var lastCell = cells.LastCell;
        int column;
        if (lastCell == null)
        {
            column = 0;
        }
        else
        {
            column = lastCell.Column + 1;
        }

        var cell = cells[0, column];
        cell.PutValue(heading);
        var style = cell.GetStyle();
        style.VerticalAlignment = TextAlignmentType.Top;
        modifyStyle?.Invoke(style);

        cell.SetStyle(style);

        if (width == null)
        {
            sheet.AutoFitColumn(column);
            var columnWidth = cells.GetColumnWidth(column);
            cells.SetColumnWidth(column, columnWidth + 3);
        }
        else
        {
            cells.SetColumnWidth(column, width.Value);
        }

        return column;
    }
}