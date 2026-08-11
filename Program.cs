using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

internal static class Program
{
    // Return meaningful exit codes for automation:
    // 0 = success
    // 1 = configuration / validation error
    // 2 = file not found
    // 3 = permission denied
    // 4 = IO error
    // 5 = interop / COM error
    // 99 = unexpected error
    private static int Main()
    {
        Logger.Initialize(Path.Combine(AppContext.BaseDirectory, "logs", "mailmerge.log"));

        try
        {
            Logger.Info("Starting MailMerge");

            string configPath = Path.Combine(AppContext.BaseDirectory, "Configuration.json");
            if (!File.Exists(configPath))
                throw new FileNotFoundException("Configuration file not found.", configPath);

            Configuration configuration = ConfigurationLoader.Load(configPath)
                ?? throw new InvalidOperationException("ConfigurationLoader returned null.");

            string templatePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, configuration.Template.Path));
            string dataPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, configuration.Data.Path));
            string outputPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, configuration.Output.Path));

            // Validate inputs early
            if (!File.Exists(templatePath))
                throw new FileNotFoundException("Template file not found.", templatePath);
            if (!File.Exists(dataPath))
                throw new FileNotFoundException("Data file not found.", dataPath);
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new InvalidOperationException("Output path is not configured.");

            Logger.Info("Configuration validated", ("template", templatePath), ("data", dataPath), ("output", outputPath));

            MailMergeTable table = new ExcelImporter().Import(
                dataPath,
                configuration.Data.Sheet,
                configuration.Data.Ranges,
                configuration.Output.FilenameTemplate);

            TableWriter.Write(Path.Combine(AppContext.BaseDirectory, "Table.json"), table);

            if (table.RowCount == 0)
            {
                Logger.Info("No rows to process - exiting");
                return 0;
            }

            using WordDocumentGenerator generator = new();
            generator.Generate(
                templatePath,
                outputPath,
                table,
                (Word.WdSaveFormat)configuration.Output.WdSaveFormat);

            Logger.Info("Mail merge completed successfully");
            return 0;
        }
        catch (FileNotFoundException fnf)
        {
            Logger.Error(fnf, "Required file missing: {File}", fnf.FileName);
            return 2;
        }
        catch (UnauthorizedAccessException ua)
        {
            Logger.Error(ua, "Permission denied");
            return 3;
        }
        catch (IOException ioEx)
        {
            Logger.Error(ioEx, "I/O error while accessing files");
            return 4;
        }
        catch (System.Runtime.InteropServices.COMException comEx)
        {
            Logger.Error(comEx, "Word interop or COM error");
            return 5;
        }
        catch (InvalidOperationException inv)
        {
            Logger.Error(inv, "Invalid operation / configuration problem");
            return 1;
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Unexpected error");
            return 99;
        }
        finally
        {
            Logger.Close();
        }
    }
}
