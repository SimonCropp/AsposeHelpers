namespace Aspose.Cells;

public static partial class CellsExtensions
{
    extension(Worksheet sheet)
    {
        public Cell AppendLinkCell(int row, string link, string? text)
        {
            var cell = AppendCell(sheet, row, text ?? link);
            sheet.Hyperlinks.Add(cell.Row,cell.Column, 1, 1, link);
            return cell;
        }

        public Cell AppendCell(int row, string? value)
        {
            var cell = sheet.FirstNullCell(row);
            if (value == null)
            {
                cell.PutValue("");
                return cell;
            }

            cell.PutValue(value);
            return cell;
        }

        public Cell AppendCell(int row, IEnumerable<string> value) =>
            sheet.AppendCell(row, string.Join(", ", value));

        public Cell AppendCell(int row, Guid? value) =>
            sheet.AppendCell(row, value?.ToString());

        public Cell AppendCell(int row, int? value) =>
            sheet.AppendCell(row, value?.ToString());

        public Cell AppendCell(int row, bool? value)
        {
            var cell = sheet.FirstNullCell(row);
            if (value == null)
            {
                cell.PutValue("");
                return cell;
            }

            cell.PutValue(value);
            return cell;
        }

        public Cell AppendCellHtml(int row, string? value)
        {
            var cell = sheet.FirstNullCell(row);
            cell.SafeSetHtml(value);
            return cell;
        }
    }

    extension(Cell cell)
    {
        public void SafeSetHtml(string? value)
        {
            if (value == null)
            {
                cell.PutValue("");
                return;
            }

            try
            {
                cell.HtmlString = value;
            }
            catch
            {
                cell.Value = value;
            }
        }
    }

    extension(Worksheet sheet)
    {
        public Cell AppendCell(int row, decimal? value)
        {
            var cell = sheet.FirstNullCell(row);
            if (value == null)
            {
                cell.PutValue("");
                return cell;
            }

            cell.PutValue((double) value.Value);
            return cell;
        }

        public Cell AppendCell(int row, Date? value, string? format = null)
        {
            var cell = sheet.FirstNullCell(row);
            if (value == null)
            {
                cell.PutValue("");
                return cell;
            }

            return sheet.AppendCell(row, value.Value.ToDateTime(new(0)), format);
        }

        public Cell AppendCell(int row, DateTimeOffset? value, string? format = null)
        {
            var cell = sheet.FirstNullCell(row);
            if (value == null)
            {
                cell.PutValue("");
                return cell;
            }

            return sheet.AppendCell(row, value.Value.DateTime, format);
        }

        public Cell AppendCell(int row, DateTime? value, string? format = null)
        {
            var cell = sheet.FirstNullCell(row);
            if (value == null)
            {
                cell.PutValue("");
                return cell;
            }

            var dateTime = value.Value;
            if (dateTime == DateTime.MinValue)
            {
                cell.PutValue("");
                return cell;
            }

            cell.PutValue(dateTime);
            var style = cell.GetStyle();
            style.Custom = GetFormat(dateTime, format);

            cell.SetStyle(style);
            return cell;
        }
    }

    static string GetFormat(DateTime value, string? format)
    {
        if (format != null)
        {
            return format;
        }

        if (value.TimeOfDay == TimeSpan.Zero)
        {
            return "yyyy-MM-dd";
        }

        if (value is {Second: 0, Millisecond: 0})
        {
            return "yyyy-MM-dd HH:mm";
        }

        if (value.Millisecond == 0)
        {
            return "yyyy-MM-dd HH:mm:ss";
        }

        return "yyyy-MM-dd HH:mm:ss.FFFFFFF";
    }
}