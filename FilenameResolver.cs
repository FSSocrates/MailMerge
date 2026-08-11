public static class FilenameResolver
{
    public static string Resolve(
        string[] template,
        string[] placeholders,
        string[][] data,
        int row)
    {
        List<string> result = [];

        for (int i = 0; i < template.Length; i++)
        {
            if (i % 2 == 0)
            {
                result.Add(template[i]);
                continue;
            }

            int index = Array.IndexOf(
                placeholders,
                template[i]);

            if (index < 0)
                throw new InvalidOperationException(
                    $"Unknown placeholder: {template[i]}");

            result.Add(data[index][row]);
        }

        return string.Concat(result);
    }
}