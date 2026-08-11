using System.Text.Json;

public static class TableWriter
{
    public static void Write(
        string path,
        MailMergeTable table)
    {
        var output = new
        {
            filename = table.Filenames,
            placeholder = table.Placeholders,
            data = table.Data
        };

        JsonSerializerOptions options = new()
        {
            WriteIndented = true
        };

        File.WriteAllText(
            path,
            JsonSerializer.Serialize(
                output,
                options));
    }
}
