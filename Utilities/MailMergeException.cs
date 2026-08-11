using System;

internal class MailMergeException : Exception
{
    public int ExitCode { get; }

    public MailMergeException(string message, int exitCode = 1, Exception? inner = null)
        : base(message, inner)
    {
        ExitCode = exitCode;
    }
}
