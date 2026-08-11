using Word = Microsoft.Office.Interop.Word;

public sealed class WordDocumentGenerator : IDisposable
{
    private readonly Word.Application application;

    public WordDocumentGenerator()
    {
        application = new Word.Application
        {
            Visible = false,
            DisplayAlerts =
                Word.WdAlertLevel.wdAlertsNone
        };
    }

    public void Generate(
        string templatePath,
        string outputDirectory,
        MailMergeTable table,
        Word.WdSaveFormat saveFormat)
    {
        Directory.CreateDirectory(outputDirectory);

        for (int row = 0; row < table.RowCount; row++)
        {
            GenerateOne(
                templatePath,
                outputDirectory,
                table,
                row,
                saveFormat);
        }
    }

    private void GenerateOne(
        string templatePath,
        string outputDirectory,
        MailMergeTable table,
        int row,
        Word.WdSaveFormat saveFormat)
    {
        Word.Document? document = null;

        string outputPath = Path.Combine(
            outputDirectory,
            table.Filenames[row]);

        try
        {
            document = application.Documents.Open(
                templatePath,
                ReadOnly: false,
                Visible: false);

            foreach (string placeholder
                     in table.Placeholders)
            {
                WordStoryReplacer.Replace(
                    document,
                    "{{" + placeholder + "}}",
                    table.GetValue(
                        placeholder,
                        row));
            }

            if (!IsValid(document))
                return;

            object fileName = outputPath;
            object fileFormat = saveFormat;

            document.SaveAs2(
                ref fileName,
                ref fileFormat);
        }
        catch
        {
            TryDelete(outputPath);
        }
        finally
        {
            document?.Close(
                Word.WdSaveOptions.wdDoNotSaveChanges);

            Com.Release(document);
        }
    }

    private static bool IsValid(
        Word.Document document)
    {
        // A document is valid if no {{...}} placeholders remain.
        return !WordPlaceholderFinder.Contains(
            document);
    }

    public void Dispose()
    {
        application.Quit(
            Word.WdSaveOptions.wdDoNotSaveChanges);

        Com.Release(application);
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
        }
    }
}