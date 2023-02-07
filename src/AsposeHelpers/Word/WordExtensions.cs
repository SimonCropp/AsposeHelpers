﻿using System.Drawing;

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
}