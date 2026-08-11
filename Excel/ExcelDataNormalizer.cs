using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

public static class ExcelDataNormalizer
{
    public static string[] FlattenTextArray(
        this Excel.Range range)
    {
        if (range == null)
            return [];

        range.Copy();

        string text;

        try
        {
            text = Clipboard.GetText(
                TextDataFormat.UnicodeText);
        }
        finally
        {
            Clipboard.Clear();
        }

        if (string.IsNullOrEmpty(text))
            return [];

        text = text.TrimEnd('\r', '\n');

        if (text.Length == 0)
            return [];

        return text
            .Split('\n')
            .SelectMany(row =>
                row.TrimEnd('\r').Split('\t'))
            .ToArray();
    }

    public static string[][] NormalizeArrays(
        Excel.Range[] ranges)
    {
        if (ranges.Length == 0)
            return [];

        string[][] arrays =
            ranges.Select(
                range => range.FlattenTextArray())
            .ToArray();

        int max = arrays.Max(
            array => array.Length);

        if (max == 0)
            return arrays;

        return arrays
            .Select(array =>
            {
                if (array.Length == 0)
                    return Enumerable
                        .Repeat("", max)
                        .ToArray();

                return Enumerable
                    .Range(0, max)
                    .Select(i =>
                        array[i % array.Length])
                    .ToArray();
            })
            .ToArray();
    }
}
