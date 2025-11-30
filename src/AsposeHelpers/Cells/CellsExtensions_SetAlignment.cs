namespace Aspose.Cells;

public static partial class CellsExtensions
{
    extension(Cell cell)
    {
        public void AlignLeft() =>
            cell.SetHorizontalAlignment(TextAlignmentType.Left);

        public void AlignRight() =>
            cell.SetHorizontalAlignment(TextAlignmentType.Right);

        public void AlignBottom() =>
            cell.SetVerticalAlignment(TextAlignmentType.Bottom);

        public void AlignTop() =>
            cell.SetVerticalAlignment(TextAlignmentType.Top);

        public void SetHorizontalAlignment(TextAlignmentType alignment)
        {
            var style = cell.GetStyle();
            style.HorizontalAlignment = alignment;
            cell.SetStyle(style);
        }

        public void SetVerticalAlignment(TextAlignmentType alignment)
        {
            var style = cell.GetStyle();
            style.VerticalAlignment = alignment;
            cell.SetStyle(style);
        }
    }

    extension(Worksheet sheet)
    {
        public void AlignLeft() =>
            sheet.SetHorizontalAlignment(TextAlignmentType.Left);

        public void AlignRight() =>
            sheet.SetHorizontalAlignment(TextAlignmentType.Right);

        public void AlignBottom() =>
            sheet.SetVerticalAlignment(TextAlignmentType.Bottom);

        public void AlignTop() =>
            sheet.SetVerticalAlignment(TextAlignmentType.Top);

        public void SetHorizontalAlignment(TextAlignmentType alignment)
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

        public void SetVerticalAlignment(TextAlignmentType alignment)
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
}