using System.Runtime.InteropServices;

public static class Com
{
    public static void Release(object? value)
    {
        if (value != null &&
            Marshal.IsComObject(value))
        {
            Marshal.FinalReleaseComObject(value);
        }
    }
}