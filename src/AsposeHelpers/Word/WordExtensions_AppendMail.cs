using Aspose.Email;

namespace Aspose.Words;

public static partial class WordExtensions
{
    public static void AppendMail(this DocumentBuilder builder, Stream stream)
    {
        using var mail = MailMessage.Load(stream);
        AppendMail(builder, mail);
    }

    public static void AppendMail(this DocumentBuilder builder, string file)
    {
        using var mail = MailMessage.Load(file);
        AppendMail(builder, mail);
    }

    public static void AppendMail(this DocumentBuilder builder, MailMessage mail)
    {
        builder.WriteH3("Email:");
        if (mail.Subject != null)
        {
            builder.Writeln($"Subject: {mail.Subject}");
        }

        if (mail.From != null)
        {
            builder.Writeln($"From: {mail.From}");
        }

        if (mail.To.Any())
        {
            builder.Writeln($"To: {mail.To}");
        }

        if (mail.CC.Any())
        {
            builder.Writeln($"To: {mail.CC}");
        }

        if (mail.Bcc.Any())
        {
            builder.Writeln($"To: {mail.Bcc}");
        }

        using var htmlStream = new MemoryStream();
        mail.Save(htmlStream, SaveOptions.DefaultHtml);
        builder.InsertHtml(Encoding.UTF8.GetString(htmlStream.GetBuffer()));
    }
}