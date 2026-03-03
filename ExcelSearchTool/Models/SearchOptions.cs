namespace ExcelSearchTool.Models;

public sealed class SearchOptions
{
    public required string Query { get; init; }
    public bool SearchContent { get; init; } = true;
    public bool SearchFileName { get; init; } = true;
    public bool UseRegex { get; init; }
    public bool CaseSensitive { get; init; }
    public required IReadOnlyCollection<string> Folders { get; init; }
}
