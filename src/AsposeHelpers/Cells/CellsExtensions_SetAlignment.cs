namespace Aspose.Cells;

public static partial class CellsExtensions
{
    public static void AlignLeft(this Cell cell) =>
        cell.SetHorizontalAlignment(TextAlignmentType.Left);

    public static void AlignRight(this Cell cell) =>
        cell.SetHorizontalAlignment(TextAlignmentType.Right);

    public static void AlignBottom(this Cell cell) =>
        cell.SetVerticalAlignment(TextAlignmentType.Bottom);

    public static void AlignTop(this Cell cell) =>
        cell.SetVerticalAlignment(TextAlignmentType.Top);

    public static void SetHorizontalAlignment(this Cell cell, TextAlignmentType alignment)
    {
        var style = cell.GetStyle();
        style.HorizontalAlignment = alignment;
        cell.SetStyle(style);
    }

    public static void SetVerticalAlignment(this Cell cell, TextAlignmentType alignment)
    {
        var style = cell.GetStyle();
        style.VerticalAlignment = alignment;
        cell.SetStyle(style);
    }

    public static void AlignLeft(this Worksheet sheet) =>
        sheet.SetHorizontalAlignment(TextAlignmentType.Left);

    public static void AlignRight(this Worksheet sheet) =>
        sheet.SetHorizontalAlignment(TextAlignmentType.Right);

    public static void AlignBottom(this Worksheet sheet) =>
        sheet.SetVerticalAlignment(TextAlignmentType.Bottom);

    public static void AlignTop(this Worksheet sheet) =>
        sheet.SetVerticalAlignment(TextAlignmentType.Top);

    public static void SetHorizontalAlignment(this Worksheet sheet, TextAlignmentType alignment)
    {
        var cells = sheet.Cells;
        var lastCell = cells.LastCell;
        for (var column = 0; column <= lastCell.Column; column++)
        {
            for (var row = 0; row <= lastCell.Row; row++)
            {
                var cell = cells[row, column];
                cell.SetHorizontalAlignment(alignment);
            }
        }
    }

    public static void SetVerticalAlignment(this Worksheet sheet, TextAlignmentType alignment)
    {
        var cells = sheet.Cells;
        var lastCell = cells.LastCell;
        for (var column = 0; column <= lastCell.Column; column++)
        {
            for (var row = 0; row <= lastCell.Row; row++)
            {
                var cell = cells[row, column];
                cell.SetVerticalAlignment(alignment);
            }
        }
    }
}