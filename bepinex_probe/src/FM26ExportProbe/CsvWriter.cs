using System.Collections.Generic;
using System.Text;

namespace FM26ExportProbe;

/// <summary>
/// Generic CSV escaping helper. No FM26-specific assumptions.
/// </summary>
public static class CsvWriter
{
    public static string Escape(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return "";
        }

        var needsQuotes = value.Contains(',')
            || value.Contains('"')
            || value.Contains('\n')
            || value.Contains('\r');

        if (!needsQuotes)
        {
            return value;
        }

        return $"\"{value.Replace("\"", "\"\"")}\"";
    }

    public static string JoinRow(IEnumerable<string?> cells)
    {
        var sb = new StringBuilder();
        var first = true;
        foreach (var cell in cells)
        {
            if (!first)
            {
                sb.Append(',');
            }
            sb.Append(Escape(cell));
            first = false;
        }
        return sb.ToString();
    }

    public static string Build(IEnumerable<string> headers, IEnumerable<IEnumerable<string?>> rows)
    {
        var sb = new StringBuilder();
        sb.AppendLine(JoinRow(headers));
        foreach (var row in rows)
        {
            sb.AppendLine(JoinRow(row));
        }
        return sb.ToString();
    }
}
