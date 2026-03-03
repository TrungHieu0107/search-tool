namespace ExcelSearchTool.Models;

public sealed class SearchResult
{
    public required string FileName { get; init; }
    public required string FilePath { get; init; }
    public required string FileType { get; init; }
    public required string Location { get; init; }
    public required string Content { get; init; }
}
