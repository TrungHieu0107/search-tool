namespace ExcelSearchTool.Models;

public sealed class SearchResult
{
    public required string FilePath { get; init; }
    public required string FileName { get; init; }
    public string MatchType { get; init; } = string.Empty;
    public string MatchDetail { get; init; } = string.Empty;
    public int? SheetIndex { get; init; }
    public string? SheetName { get; init; }
    public int? Row { get; init; }
    public int? Column { get; init; }
}
