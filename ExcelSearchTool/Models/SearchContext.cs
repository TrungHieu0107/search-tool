using System.Collections.Concurrent;
using System.Text.RegularExpressions;

namespace ExcelSearchTool.Models;

public sealed class SearchContext
{
    public required SearchOptions Options { get; init; }
    public required Regex? Regex { get; init; }
    public required CancellationToken CancellationToken { get; init; }
    public required ConcurrentBag<SearchResult> Results { get; init; }

    public bool IsMatch(string value)
    {
        return Regex is not null
            ? Regex.IsMatch(value)
            : value.Contains(Options.Query, Options.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
    }

    public string HighlightMatch(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return value;
        }

        if (Regex is not null)
        {
            try
            {
                return Regex.Replace(value, m => $"[[{m.Value}]]", 1);
            }
            catch
            {
                return value;
            }
        }

        var comparison = Options.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var index = value.IndexOf(Options.Query, comparison);
        if (index < 0)
        {
            return value;
        }

        return string.Concat(
            value[..index],
            "[[",
            value.Substring(index, Options.Query.Length),
            "]]",
            value[(index + Options.Query.Length)..]);
    }
}
