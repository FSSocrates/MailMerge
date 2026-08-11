public static class SaveFormatExtensions
{
    public static readonly string[] Values =
    {
        ".doc",
        ".dot",
        ".txt",
        ".txt",
        ".txt",
        ".txt",
        ".rtf",
        ".txt",
        ".html",
        ".mhtml",
        ".html",
        ".xml",
        ".docx",
        ".docm",
        ".dotx",
        ".dotm",
        ".docx",
        ".pdf",
        ".xps",
        ".xml",
        ".xml",
        ".xml",
        ".xml",
        ".odt"
    };

    public static string GetExtension(int format)
    {
        if ((uint)format >= Values.Length)
            throw new ArgumentOutOfRangeException(
                nameof(format));

        return Values[format];
    }
}
