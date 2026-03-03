using System.Collections.Concurrent;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using ExcelSearchTool.Models;
using OfficeOpenXml;

namespace ExcelSearchTool.Services;

public sealed class ContentSearchService(ErrorLogger errorLogger)
{
    private readonly ErrorLogger _errorLogger = errorLogger;

    public bool CanContainMatch(string filePath, SearchOptions options)
    {
        if (!options.SearchContent)
        {
            return false;
        }

        if (options.UseRegex)
        {
            return true;
        }

        try
        {
            using var archive = ZipFile.OpenRead(filePath);
            foreach (var entry in archive.Entries)
            {
                if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                using var stream = entry.Open();
                using var reader = new StreamReader(stream, Encoding.UTF8, true, 1024, false);
                var xml = reader.ReadToEnd();
                if (xml.Contains(options.Query, options.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }
        catch
        {
            return true;
        }

        return false;
    }

    public async Task<IReadOnlyCollection<SearchResult>> FindMatchesAsync(string filePath, SearchOptions options, Regex? regex, CancellationToken cancellationToken)
    {
        var results = new ConcurrentBag<SearchResult>();

        try
        {
            await using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var package = new ExcelPackage(stream);
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (worksheet.Dimension is null)
                {
                    continue;
                }

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;
                for (var row = start.Row; row <= end.Row; row++)
                {
                    for (var column = start.Column; column <= end.Column; column++)
                    {
                        var value = worksheet.Cells[row, column].Text;
                        if (string.IsNullOrEmpty(value))
                        {
                            continue;
                        }

                        var isMatch = regex is not null
                            ? regex.IsMatch(value)
                            : value.Contains(options.Query, options.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);

                        if (!isMatch)
                        {
                            continue;
                        }

                        results.Add(new SearchResult
                        {
                            FilePath = filePath,
                            FileName = Path.GetFileName(filePath),
                            MatchType = "Content",
                            MatchDetail = value.Length > 90 ? string.Concat(value.AsSpan(0, 90), "...") : value,
                            SheetName = worksheet.Name,
                            SheetIndex = worksheet.Index,
                            Row = row,
                            Column = column
                        });
                    }
                }
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            await _errorLogger.LogAsync($"Skipped file due to read/parse error: {filePath}", ex).ConfigureAwait(false);
        }

        return results.ToArray();
    }
}
