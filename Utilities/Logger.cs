using System;
using System.IO;
using System.Threading;

internal static class Logger
{
    private static readonly object _lock = new();
    private static string? _path;

    public static void Initialize(string path)
    {
        _path = path;
        try
        {
            var dir = Path.GetDirectoryName(_path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            // Write header
            WriteLine("----- MailMerge Log Started: " + DateTimeOffset.Now + " -----");
        }
        catch
        {
            // If logging initialization fails, swallow so it doesn't block the program start.
            _path = null;
        }
    }

    public static void Info(string message, params (string key, object? value)[] props)
    {
        WriteLine($"INFO  {DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} {message}{FormatProps(props)}");
    }

    public static void Error(Exception ex, string message = "", params object[] formatArgs)
    {
        try
        {
            var msg = string.IsNullOrEmpty(message) ? ex.Message : string.Format(message, formatArgs);
            WriteLine($"ERROR {DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} {msg}");
            WriteLine(ex.ToString());
        }
        catch
        {
            // Swallow
        }
    }

    private static string FormatProps((string key, object? value)[] props)
    {
        if (props == null || props.Length == 0) return string.Empty;
        try
        {
            return " | " + string.Join(", ", Array.ConvertAll(props, p => $"{p.key}={p.value}"));
        }
        catch { return string.Empty; }
    }

    private static void WriteLine(string line)
    {
        if (string.IsNullOrEmpty(_path)) return;
        try
        {
            lock (_lock)
            {
                File.AppendAllText(_path, line + Environment.NewLine);
            }
        }
        catch
        {
            // Logging must not throw.
        }
    }

    public static void Close()
    {
        // Nothing to dispose with this simple implementation, but keep method for future.
        WriteLine("----- MailMerge Log Ended: " + DateTimeOffset.Now + " -----");
    }
}
