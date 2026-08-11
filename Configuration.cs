public sealed class Configuration
{
    public TemplateConfiguration Template { get; set; } = new();
    public OutputConfiguration Output { get; set; } = new();
    public DataConfiguration Data { get; set; } = new();
}

public sealed class TemplateConfiguration
{
    public string Path { get; set; } = "";
}

public sealed class OutputConfiguration
{
    public string Path { get; set; } = "";
    public string[] FilenameTemplate { get; set; } = [];
    public int WdSaveFormat { get; set; }
}

public sealed class DataConfiguration
{
    public string Path { get; set; } = "";
    public string Sheet { get; set; } = "";
    public Dictionary<string, string> Ranges { get; set; } = [];
}