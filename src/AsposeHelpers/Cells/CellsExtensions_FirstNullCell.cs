namespace Aspose.Cells;

public static partial class CellsExtensions
{
    extension(Worksheet sheet)
    {
        public Cell FirstNullCell(int row)
        {
            var column = 0;
            while (true)
            {
                var cell = sheet.Cells[row, column];
                if (cell.Value == null)
                {
                    return cell;
                }

                column++;
            }
        }
    }
}