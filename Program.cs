using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
internal static class Program {
    private static int Main() {
        try {
            string configPath = Path.Combine(AppContext.BaseDirectory, "Configuration.json");
            if(!File.Exists(configPath)) {
                Log("✖", "Configuration.json [File not found]");
                return 2;
            }
            Configuration configuration = ConfigurationLoader.Load(configPath);
            Log("✔", "Configuration.json");
            string templatePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, configuration.Template.Path));
            string dataPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, configuration.Data.Path));
            string outputPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, configuration.Output.Path));
            if(!File.Exists(templatePath)) {
                Log("✖", $ "{Path.GetFileName(templatePath)} [File not found]");
                return 2;
            }
            if(!File.Exists(dataPath)) {
                Log("✖", $ "{Path.GetFileName(dataPath)} [File not found]");
                return 2;
            }
            MailMergeTable table = new ExcelImporter().Import(dataPath, configuration.Data.Sheet, configuration.Data.Ranges, configuration.Output.FilenameTemplate);
            Log("✔", Path.GetFileName(dataPath));
            string tablePath = Path.Combine(AppContext.BaseDirectory, "Table.json");
            TableWriter.Write(tablePath, table);
            Log("✔", "Table.json");
            if(table.RowCount == 0) {
                Log("⚠", "Completed — no documents generated");
                return 0;
            }
            int generated = 0;
            int failed = 0;
            using WordDocumentGenerator generator = new();
            generator.Generate(templatePath, outputPath, table,
                (fileName, error) => {
                    if(error is null) {
                        generated++;
                        Log("✔", fileName);
                    } else {
                        failed++;
                        Log("✖", $ "{fileName} [{error}]");
                    }
                },
                (Word.WdSaveFormat) configuration.Output.WdSaveFormat);
            string status = failed == 0 ? "✔" : "⚠";
            Log(status, $ "Completed — {generated} generated, {failed} failed");
            return failed == 0 ? 0 : 1;
        } catch (FileNotFoundException ex) {
            Log("✖", $ "{Path.GetFileName(ex.FileName)} [File not found]");
            return 2;
        } catch (UnauthorizedAccessException) {
            Log("✖", "Access denied");
            return 3;
        } catch (IOException ex) {
            Log("✖", $ "I/O error [{ex.Message}]");
            return 4;
        } catch (System.Runtime.InteropServices.COMException ex) {
            Log("✖", $ "Office error [{ex.Message}]");
            return 5;
        } catch (InvalidOperationException ex) {
            Log("✖", ex.Message);
            return 1;
        } catch (Exception ex) {
            Log("✖", ex.Message);
            return 99;
        }
    }
    private static void Log(string symbol, string message) {
        Console.WriteLine($ "[{DateTime.Now:HH:mm:ss}] {symbol} {message}");
    }
}
