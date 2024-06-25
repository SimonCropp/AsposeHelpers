using System.Diagnostics.Contracts;
using System.Drawing;
using Aspose.Words.Fields;

namespace Aspose.Words;

public static partial class WordExtensions
{
    public static void WriteEmail(this DocumentBuilder builder, string email)
    {
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;
        builder.InsertHyperlink(email, $"mailto://{email}", false);
        builder.Font.ClearFormatting();
    }

    public static void ClearFormatting(this DocumentBuilder builder)
    {
        builder.Bold = false;
        builder.Italic = false;
        builder.Font.ClearFormatting();
        builder.ParagraphFormat.ClearFormatting();
        builder.Font.Border.ClearFormatting();
    }

    public static void WriteLine(this DocumentBuilder builder, char line) =>
        builder.Writeln(line.ToString());

    public static void WriteBoldLine(this DocumentBuilder builder, string line)
    {
        builder.Bold = true;
        builder.Writeln(line);
        builder.Font.ClearFormatting();
    }

    public static void WriteItalicLine(this DocumentBuilder builder, string line)
    {
        builder.Italic = true;
        builder.Writeln(line);
        builder.Font.ClearFormatting();
    }

    public static void WriteBoldItalicLine(this DocumentBuilder builder, string line)
    {
        builder.Bold = true;
        builder.Italic = true;
        builder.Writeln(line);
        builder.Font.ClearFormatting();
    }

    public static void WriteBold(this DocumentBuilder builder, string line)
    {
        builder.Bold = true;
        builder.Write(line);
        builder.Font.ClearFormatting();
    }

    public static void WriteItalic(this DocumentBuilder builder, string line)
    {
        builder.Italic = true;
        builder.Write(line);
        builder.Font.ClearFormatting();
    }

    public static void WriteBoldItalic(this DocumentBuilder builder, string line)
    {
        builder.Bold = true;
        builder.Italic = true;
        builder.Write(line);
        builder.Font.ClearFormatting();
    }

    [Pure]
    public static IDisposable UseStyled(this DocumentBuilder builder, string name)
    {
        builder.ParagraphFormat.Style = FindStyle(builder, name);
        builder.Bold = true;
        return new ClearStyleDisposable(builder);
    }

    [Pure]
    public static IDisposable UseBold(this DocumentBuilder builder)
    {
        builder.Bold = true;
        return new FontClearFormattingDisposable(builder);
    }

    [Pure]
    public static IDisposable UseItalic(this DocumentBuilder builder)
    {
        builder.Italic = true;
        return new FontClearFormattingDisposable(builder);
    }

    [Pure]
    public static IDisposable UseBoldItalic(this DocumentBuilder builder)
    {
        builder.Bold = true;
        builder.Italic = true;
        return new FontClearFormattingDisposable(builder);
    }

    public static void Write(this DocumentBuilder builder, char ch) =>
        builder.Write(ch.ToString());

    public static FieldTC InsertTocEntry(this DocumentBuilder builder, string text, int level, bool pageNumber = true) =>
        InsertTocEntry(builder,  text, level.ToString(), pageNumber);

    public static FieldTC InsertTocEntry(this DocumentBuilder builder, string text, string level, bool pageNumber = true)
    {
        builder.Font.ClearFormatting();
        builder.Font.Color = Color.White;
        builder.Font.Size = 0;
        var field = (FieldTC) builder.InsertField(FieldType.FieldTOCEntry, true);
        field.EntryLevel = level;
        field.OmitPageNumber = !pageNumber;
        field.Text = text;
        builder.Writeln();
        builder.Font.ClearFormatting();
        return field;
    }

    public static void WriteLink(this DocumentBuilder builder, string text, string link)
    {
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;
        builder.InsertHyperlink(text, link, false);
        builder.Font.ClearFormatting();
    }

    public static void WriteH1(this DocumentBuilder builder, string text) =>
        builder.WriteStyled(text, StyleIdentifier.Heading1);

    public static void WriteH2(this DocumentBuilder builder, string text) =>
        builder.WriteStyled(text, StyleIdentifier.Heading2);

    public static void WriteH3(this DocumentBuilder builder, string text) =>
        builder.WriteStyled(text, StyleIdentifier.Heading3);

    public static void WriteH4(this DocumentBuilder builder, string text) =>
        builder.WriteStyled(text, StyleIdentifier.Heading4);

    public static void WriteH5(this DocumentBuilder builder, string text) =>
        builder.WriteStyled(text, StyleIdentifier.Heading5);

    public static void SetMargins(this DocumentBuilder builder, double millimeters)
    {
        var margin = ConvertUtil.MillimeterToPoint(millimeters);
        var pageSetup = builder.PageSetup;
        pageSetup.TopMargin = margin;
        pageSetup.BottomMargin = margin;
        pageSetup.LeftMargin = margin;
        pageSetup.RightMargin = margin;
        pageSetup.HeaderDistance = margin;
        pageSetup.FooterDistance = margin;
    }

    public static void WriteStyled(this DocumentBuilder builder, string text, string styleName)
    {
        var style = FindStyle(builder, styleName);

        builder.ParagraphFormat.Style = style;
        builder.Writeln(text);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
    }

    static Style? FindStyle(DocumentBuilder builder, string styleName)
    {
        var styles = builder.Document
            .Styles
            .Where(_=>_.Type == StyleType.Paragraph)
            .ToList();

        var style = styles.SingleOrDefault(_ => _.Name == styleName);
        if (style == null)
        {
            throw new($"Could not find paragraph {styleName}. Available styles: {string.Join(", ", styles.Select(_ => _.Name))}");
        }

        return style;
    }

    public static void WriteStyled(this DocumentBuilder builder, string text, StyleIdentifier style)
    {
        builder.ParagraphFormat.StyleIdentifier = style;
        builder.Writeln(text);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
    }

    public static void ApplyBorder(this DocumentBuilder documentBuilder, LineStyle style)
    {
        var borders = documentBuilder.ParagraphFormat.Borders;
        borders[BorderType.Left].LineStyle = style;
        borders[BorderType.Right].LineStyle = style;
        borders[BorderType.Top].LineStyle = style;
        borders[BorderType.Bottom].LineStyle = style;
    }

    public static void AddPageNumbers(this DocumentBuilder builder)
    {
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
        builder.Write("Page ");
        builder.InsertField(FieldType.FieldPage, true);
        builder.Write(" of ");
        builder.InsertField(FieldType.FieldNumPages, true);
    }
}