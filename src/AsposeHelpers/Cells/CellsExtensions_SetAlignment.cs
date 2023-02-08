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
}