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

            string templatePath = Path.GetFullPath(Path.Combine(
                AppContext.BaseDirectory,
                configuration.Template.Path));

            string dataPath = Path.GetFullPath(Path.Combine(
                AppContext.BaseDirectory,
                configuration.Data.Path));

            string outputPath = Path.GetFullPath(Path.Combine(
                AppContext.BaseDirectory,
                configuration.Output.Path));

            MailMergeTable table =
                new ExcelImporter().Import(
                    dataPath,
                    configuration.Data.Sheet,
                    configuration.Data.Ranges,
                    configuration.Output.FilenameTemplate);

            TableWriter.Write(
                Path.Combine(AppContext.BaseDirectory, "Table.json"),
                table);

            if (table.RowCount == 0)
                return;

            using WordDocumentGenerator generator = new();

            generator.Generate(
                templatePath,
                outputPath,
                table,
                (Word.WdSaveFormat)
                    configuration.Output.WdSaveFormat);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            Environment.ExitCode = 1;
        }
    }
}
