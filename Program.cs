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

            new MergeCoordinator(
                configuration,
                configPath).Run();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
        }
    }
}
