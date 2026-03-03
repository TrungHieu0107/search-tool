namespace ExcelSearchTool.Models;

public sealed class SearchOptions
{
    public required string Query { get; init; }
    public bool UseRegex { get; init; }
    public bool CaseSensitive { get; init; }
    public bool SearchFileName { get; init; } = true;

    public bool IncludeExcel { get; init; } = true;
    public bool IncludeText { get; init; } = true;
    public bool IncludeCode { get; init; } = true;

    public long MaxFileSizeBytes { get; init; } = 10 * 1024 * 1024;
    public required IReadOnlyCollection<string> Folders { get; init; }
}
