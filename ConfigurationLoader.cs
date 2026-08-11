using System.Text.Json;

public static class ConfigurationLoader
{
    public static Configuration Load(
        string path)
    {
        string json =
            File.ReadAllText(path);

        Configuration configuration =
            JsonSerializer.Deserialize<Configuration>(
                json,
                new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                })
            ?? throw new InvalidOperationException(
                "Invalid configuration.");

        Validate(configuration);

        return configuration;
    }

    private static void Validate(
        Configuration c)
    {
        if (string.IsNullOrWhiteSpace(c.Template.Path))
            throw new InvalidOperationException(
                "template.path is required.");

        if (string.IsNullOrWhiteSpace(c.Output.Path))
            throw new InvalidOperationException(
                "output.path is required.");

        if (c.Output.FilenameTemplate.Length == 0 ||
            c.Output.FilenameTemplate.Length % 2 != 0)
        {
            throw new InvalidOperationException(
                "output.filenameTemplate must contain an even number of elements.");
        }

        if (c.Output.WdSaveFormat < 0 ||
            c.Output.WdSaveFormat >=
            SaveFormatExtensions.Values.Length)
        {
            throw new InvalidOperationException(
                "Invalid output.wdSaveFormat.");
        }

        foreach (string placeholder
                 in c.Output.FilenameTemplate
                     .Where((_, i) => i % 2 == 1))
        {
            if (!c.Data.Ranges.ContainsKey(
                    placeholder))
            {
                throw new InvalidOperationException(
                    $"Unknown filename placeholder '{placeholder}'.");
            }
        }
    }
}