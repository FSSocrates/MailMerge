using Word = Microsoft.Office.Interop.Word;

public sealed class WordDocumentGenerator : IDisposable
{
    private readonly Word.Application application;

    public WordDocumentGenerator()
    {
        application = new Word.Application
        {
            Visible = false,
            DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
        };
    }

    public void Generate(string templatePath, string outputDirectory, MailMergeTable table,
                         Action<string, string?> report, Word.WdSaveFormat saveFormat)
    {
        Directory.CreateDirectory(outputDirectory);
        for (int row = 0; row < table.RowCount; row++)
        {
            GenerateOne(templatePath, outputDirectory, table, row, report, saveFormat);
        }
    }

    private void GenerateOne(string templatePath, string outputDirectory, MailMergeTable table,
                             int row, Action<string, string?> report,
                             Word.WdSaveFormat saveFormat)
    {
        Word.Document? document = null;
        string fileName = table.Filenames[row];
        string outputPath = Path.Combine(outputDirectory, fileName);
        try
        {
            document = application.Documents.Open(templatePath, ReadOnly: false, Visible: false);

            // Replace all placeholders enclosed in {{...}} using GetValue2 for safe insertion.
            foreach (string placeholder in table.Placeholders)
            {
                string token = "{{" + placeholder + "}}";
                string safeValue = table.GetValue2(placeholder, row);
                WordStoryReplacer.Replace(document, token, safeValue);
            }

            if (!IsValid(document))
            {
                report(fileName, "Unresolved placeholder");
                return;
            }

            object fileNameObject = outputPath;
            object fileFormat = saveFormat;
            document.SaveAs2(ref fileNameObject, ref fileFormat);
            report(fileName, null);
        }
        catch (Exception ex)
        {
            report(fileName, ex.Message);
        }
        finally
        {
            document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Com.Release(document);
        }
    }

    private static bool IsValid(Word.Document document)
    {
        return !WordPlaceholderFinder.Contains(document);
    }

    public void Dispose()
    {
        application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
        Com.Release(application);
    }
}
