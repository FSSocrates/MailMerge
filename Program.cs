using Word = Microsoft.Office.Interop.Word;

internal static class Program
{
    private static void Main()
    {
        try
        {
            string configPath = Path.Combine(
                AppContext.BaseDirectory,
                "Configuration.json");

            Configuration configuration =
                ConfigurationLoader.Load(configPath);

            new Coordinator(
                configuration,
                configPath).Run();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
        }
    }
}

public sealed class Coordinator
{
    private readonly Configuration configuration;
    private readonly string baseDirectory;

    public Coordinator(
        Configuration configuration,
        string configPath)
    {
        this.configuration = configuration;

        baseDirectory = Path.GetDirectoryName(
            Path.GetFullPath(configPath))!;
    }

    public void Run()
    {
        string templatePath =
            Resolve(configuration.Template.Path);

        string dataPath =
            Resolve(configuration.Data.Path);

        string outputPath =
            Resolve(configuration.Output.Path);

        ExcelImporter importer = new();

        MailMergeTable table =
            importer.Import(
                dataPath,
                configuration.Data.Sheet,
                configuration.Data.Ranges,
                configuration.Output.FilenameTemplate);

        TableWriter.Write(
            Path.Combine(
                baseDirectory,
                "Table.json"),
            table);

        if (table.RowCount == 0)
            return;

        using WordDocumentGenerator generator = new();

        generator.Generate(
            templatePath,
            outputPath,
            table,
            (Microsoft.Office.Interop.Word.WdSaveFormat)
                configuration.Output.WdSaveFormat);
    }

    private string Resolve(string path)
    {
        return Path.GetFullPath(
            Path.Combine(baseDirectory, path));
    }
}
