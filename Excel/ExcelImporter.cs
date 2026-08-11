using Excel = Microsoft.Office.Interop.Excel;

public sealed class ExcelImporter
{
    public MailMergeTable Import(
        string workbookPath,
        string sheetName,
        Dictionary<string, string> rangeDefinitions,
        string[] filenameTemplate)
    {
        Excel.Application? app = null;
        Excel.Workbook? workbook = null;
        Excel.Worksheet? worksheet = null;

        try
        {
            app = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };

            workbook = app.Workbooks.Open(
                workbookPath,
                ReadOnly: true);

            worksheet = workbook.Worksheets[sheetName];

            string[] placeholders =
                rangeDefinitions.Keys.ToArray();

            Excel.Range[] ranges =
                placeholders
                    .Select(p =>
                        worksheet.Range[
                            rangeDefinitions[p]])
                    .ToArray();

            try
            {
                string[][] data =
                    ExcelDataNormalizer.NormalizeArrays(
                        ranges);

                int rowCount =
                    data.Length == 0
                        ? 0
                        : data.Max(x => x.Length);

                string[] filenames =
                    Enumerable.Range(0, rowCount)
                        .Select(row =>
                            FilenameResolver.Resolve(
                                filenameTemplate,
                                placeholders,
                                data,
                                row))
                        .ToArray();

                return new MailMergeTable
                {
                    Filenames = filenames,
                    Placeholders = placeholders,
                    Data = data
                };
            }
            finally
            {
                foreach (Excel.Range range in ranges)
                    Com.Release(range);
            }
        }
        finally
        {
            workbook?.Close(false);

            Com.Release(worksheet);
            Com.Release(workbook);

            app?.Quit();
            Com.Release(app);
        }
    }
}
