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
}