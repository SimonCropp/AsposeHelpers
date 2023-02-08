namespace Aspose.Cells;

public static partial class CellsExtensions
{
    public static Cell AppendCell(this Worksheet sheet, int row, string? value)
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

    public static Cell AppendCell(this Worksheet sheet, int row, Guid? value) =>
        sheet.AppendCell(row, value?.ToString());

    public static Cell AppendCell(this Worksheet sheet, int row, int? value) =>
        sheet.AppendCell(row, value?.ToString());
    public static Cell AppendCell(this Worksheet sheet, int row, bool? value)
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

    public static Cell AppendCellHtml(this Worksheet sheet, int row, string? value)
    {
        var cell = sheet.FirstNullCell(row);
        if (value == null)
        {
            cell.PutValue("");
            return cell;
        }

        try
        {
            cell.HtmlString = value;
            return cell;
        }
        catch (Exception exception)
        {
            throw new($"Unable to set html. Html: {value}", exception);
        }
    }

    public static Cell AppendCell(this Worksheet sheet, int row, decimal? value)
    {
        var cell = sheet.FirstNullCell(row);
        if (value == null)
        {
            cell.PutValue("");
            return cell;
        }

        cell.PutValue((double)value.Value);
        return cell;
    }

    public static Cell AppendCell(this Worksheet sheet, int row, Date? value)
    {
        var cell = sheet.FirstNullCell(row);
        if (value == null)
        {
            cell.PutValue("");
            return cell;
        }

        cell.PutValue(value.Value.ToDateTime(new(0)));
        var style = cell.GetStyle();
        style.Custom = "yyyy-MM-dd";
        cell.SetStyle(style);
        return cell;
    }

    public static Cell AppendCell(this Worksheet sheet, int row, DateTimeOffset? value)
    {
        var cell = sheet.FirstNullCell(row);
        if (value == null)
        {
            cell.PutValue("");
            return cell;
        }

        cell.PutValue(value.Value.DateTime);
        var style = cell.GetStyle();
        style.Custom = GetFormat(value.Value.DateTime);
        cell.SetStyle(style);
        return cell;
    }

    public static Cell AppendCell(this Worksheet sheet, int row, DateTime? value)
    {
        var cell = sheet.FirstNullCell(row);
        if (value == null)
        {
            cell.PutValue("");
            return cell;
        }

        var dateTime = value.Value;
        cell.PutValue(dateTime);
        var style = cell.GetStyle();
        style.Custom = GetFormat(dateTime);

        cell.SetStyle(style);
        return cell;
    }

    static string GetFormat(DateTime value)
    {
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