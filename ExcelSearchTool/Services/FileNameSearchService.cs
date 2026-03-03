using System.Text.RegularExpressions;
using ExcelSearchTool.Models;

namespace ExcelSearchTool.Services;

public sealed class FileNameSearchService
{
    public bool IsMatch(string filePath, SearchOptions options, Regex? regex)
    {
        if (!options.SearchFileName)
        {
            return false;
        }

        var fileName = Path.GetFileName(filePath);
        return regex is not null
            ? regex.IsMatch(fileName)
            : fileName.Contains(options.Query, options.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
    }

    public SearchResult ToResult(string filePath) => new()
    {
        FilePath = filePath,
        FileName = Path.GetFileName(filePath),
        MatchType = "FileName",
        MatchDetail = "Matched filename"
    };
}
