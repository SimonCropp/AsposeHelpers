using Aspose.Words.Fields;
using Aspose.Words.Tables;

namespace Aspose.Words;

public static partial class WordExtensions
{
    extension(DocumentBuilder builder)
    {
        public void ReplaceField(string name, string value)
        {
            var fields = builder.FindFields(name);

            foreach (var field in fields)
            {
                builder.MoveToBookmark(name, false, true);
                field.RemoveField();
                builder.Write(value);
            }
        }

        public List<FormField> FindFields(string name)
        {
            var fields = builder.Document.Range.FormFields;
            if (fields.Count == 0)
            {
                throw new($"Could not find field: {name}. Document contains no fields.");
            }

            var found = fields.Where(_ => _.Name == name || _.Result == name).ToList();
            if (found.Count == 0)
            {
                throw new(
                    $"""
                     Could not find field: {name}.
                     Existing fields are:
                     {string.Join('\n', fields.Select(_ => _.Name).Distinct().Order().Select(_ => $" * {_}"))}
                     """);
            }

            return found;
        }

        public void DisplaceField(string name)
        {
            var field = builder.FindFields(name).Single();
            builder.MoveToBookmark(name, false, true);
            field.RemoveField();
        }

        public void WriteEmail(string email)
        {
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;
            builder.InsertHyperlink(email, $"mailto://{email}", false);
            builder.Font.ClearFormatting();
        }

        public void ClearFormatting()
        {
            builder.Bold = false;
            builder.Italic = false;
            builder.Font.ClearFormatting();
            builder.ParagraphFormat.ClearFormatting();
            builder.Font.Border.ClearFormatting();
        }

        public void WriteLine(char line) =>
            builder.Writeln(line.ToString());

        public void WriteBoldLine(string line)
        {
            builder.Bold = true;
            builder.Writeln(line);
            builder.Font.ClearFormatting();
        }

        public void WriteItalicLine(string line)
        {
            builder.Italic = true;
            builder.Writeln(line);
            builder.Font.ClearFormatting();
        }

        public void WriteBoldItalicLine(string line)
        {
            builder.Bold = true;
            builder.Italic = true;
            builder.Writeln(line);
            builder.Font.ClearFormatting();
        }

        public void WriteBold(string line)
        {
            builder.Bold = true;
            builder.Write(line);
            builder.Font.ClearFormatting();
        }

        public void WriteItalic(string line)
        {
            builder.Italic = true;
            builder.Write(line);
            builder.Font.ClearFormatting();
        }

        public void WriteBoldItalic(string line)
        {
            builder.Bold = true;
            builder.Italic = true;
            builder.Write(line);
            builder.Font.ClearFormatting();
        }

        [Pure]
        public IDisposable UseStyled(string name)
        {
            builder.ParagraphFormat.Style = FindParagraphStyle(builder, name);
            return new ClearStyleDisposable(builder);
        }

        [Pure]
        public IDisposable UseBold()
        {
            builder.Bold = true;
            return new FontClearFormattingDisposable(builder);
        }

        [Pure]
        public IDisposable UseItalic()
        {
            builder.Italic = true;
            return new FontClearFormattingDisposable(builder);
        }

        [Pure]
        public IDisposable UseBoldItalic()
        {
            builder.Bold = true;
            builder.Italic = true;
            return new FontClearFormattingDisposable(builder);
        }

        public void Write(char ch) =>
            builder.Write(ch.ToString());

        public FieldTC InsertTocEntry(string text, int level, bool pageNumber = true) =>
            InsertTocEntry(builder, text, level.ToString(), pageNumber);

        public FieldTC InsertTocEntry(string text, string level, bool pageNumber = true)
        {
            builder.Font.ClearFormatting();
            builder.Font.Color = Color.White;
            builder.Font.Size = 0;
            var field = (FieldTC)builder.InsertField(FieldType.FieldTOCEntry, true);
            field.EntryLevel = level;
            field.OmitPageNumber = !pageNumber;
            field.Text = text;
            builder.Writeln();
            builder.Font.ClearFormatting();
            return field;
        }

        public void WriteLink(string text, string link)
        {
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;
            builder.InsertHyperlink(text, link, false);
            builder.Font.ClearFormatting();
        }

        public void WriteH1(string text) =>
            builder.WriteStyled(text, StyleIdentifier.Heading1);

        public void WriteH2(string text) =>
            builder.WriteStyled(text, StyleIdentifier.Heading2);

        public void WriteH3(string text) =>
            builder.WriteStyled(text, StyleIdentifier.Heading3);

        public void WriteH4(string text) =>
            builder.WriteStyled(text, StyleIdentifier.Heading4);

        public void WriteH5(string text) =>
            builder.WriteStyled(text, StyleIdentifier.Heading5);

        public void SetMargins(double millimeters)
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

        public void WriteStyled(string text, string styleName)
        {
            var style = FindParagraphStyle(builder, styleName);

            builder.ParagraphFormat.Style = style;
            builder.Writeln(text);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        }
    }

    extension(Table table)
    {
        public void AssignStyle(string name) =>
            table.Style = table.Document.FindStyle(name, StyleType.Table);
    }

    extension(DocumentBuilder builder)
    {
        public Style FindCharacterStyle(string name) =>
            FindStyle(builder
                .Document, name, StyleType.Character);
    }

    extension(DocumentBase document)
    {
        public Style FindCharacterStyle(string name) =>
            FindStyle(document, name, StyleType.Character);
    }

    extension(DocumentBuilder builder)
    {
        public Style FindTableStyle(string name) =>
            FindStyle(builder
                .Document, name, StyleType.Table);
    }

    extension(DocumentBase document)
    {
        public Style FindTableStyle(string name) =>
            FindStyle(document, name, StyleType.Table);
    }

    extension(DocumentBuilder builder)
    {
        public Style FindListStyle(string name) =>
            FindStyle(builder
                .Document, name, StyleType.List);
    }

    extension(DocumentBase document)
    {
        public Style FindListStyle(string name) =>
            FindStyle(document, name, StyleType.List);
    }

    extension(DocumentBuilder builder)
    {
        public Style FindParagraphStyle(string name) =>
            FindStyle(builder
                .Document, name, StyleType.Paragraph);
    }

    extension(DocumentBase document)
    {
        public Style FindParagraphStyle(string name) =>
            FindStyle(document, name, StyleType.Paragraph);
    }

    extension(DocumentBuilder builder)
    {
        public Style FindStyle(string name, StyleType? type = null) =>
            FindStyle(builder
                .Document, name, type);
    }

    extension(DocumentBase document)
    {
        public Style FindStyle(string name, StyleType? type)
        {
            List<Style> styles;
            var available = document
                .Styles;
            if (type == null)
            {
                styles = available
                    .ToList();
            }
            else
            {
                styles = available
                    .Where(_ => _.Type == type)
                    .ToList();
            }

            var style = styles.SingleOrDefault(_ => _.Name == name);
            if (style != null)
            {
                return style;
            }

            var names = string.Join(", ", styles.Select(_ => _.Name));
            if (type == null)
            {
                throw new($"Could not find style '{name}'. Available styles: {names}");
            }

            throw new($"Could not find {type} style '{name}'. Available styles: {names}");
        }
    }

    extension(DocumentBuilder builder)
    {
        public void WriteStyled(string text, StyleIdentifier style)
        {
            builder.ParagraphFormat.StyleIdentifier = style;
            builder.Writeln(text);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        }

        public void ApplyBorder(LineStyle style)
        {
            var borders = builder.ParagraphFormat.Borders;
            borders[BorderType.Left].LineStyle = style;
            borders[BorderType.Right].LineStyle = style;
            borders[BorderType.Top].LineStyle = style;
            borders[BorderType.Bottom].LineStyle = style;
        }

        public void AddPageNumbers()
        {
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Write("Page ");
            builder.InsertField(FieldType.FieldPage, true);
            builder.Write(" of ");
            builder.InsertField(FieldType.FieldNumPages, true);
        }
    }
}