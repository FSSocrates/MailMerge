public sealed class MailMergeTable
{
    public string[] Filenames { get; init; } = [];
    public string[] Placeholders { get; init; } = [];
    public string[][] Data { get; init; } = [];

    public int RowCount => Filenames.Length;

    public string GetValue(string placeholder, int row)
    {
        int index = Array.IndexOf(
            Placeholders,
            placeholder);

        if (index < 0)
            throw new KeyNotFoundException(
                $"Unknown placeholder: {placeholder}");

        return Data[index][row];
    }

    // GetValue2 returns a sanitized, limited and safe-to-insert string that avoids
    // accidental injection of special Word constructs. This is where we centralize
    // any value handling instead of doing ad-hoc "copy-and-paste" style manipulations.
    public string GetValue2(string placeholder, int row)
    {
        var raw = GetValue(placeholder, row) ?? string.Empty;

        // Remove embedded null characters
        raw = raw.Replace("\0", string.Empty);

        // Normalize newlines to environment newline and trim surrounding whitespace
        raw = raw.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", System.Environment.NewLine).Trim();

        // Limit length to a reasonable cap to avoid extremely large inserts
        const int MaxLength = 32_768; // 32KB
        if (raw.Length > MaxLength)
            raw = raw.Substring(0, MaxLength);

        return raw;
    }
}