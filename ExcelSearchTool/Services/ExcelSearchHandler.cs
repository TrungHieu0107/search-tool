using System.IO.Compression;
using System.Text;
using ExcelSearchTool.Models;
using OfficeOpenXml;

namespace ExcelSearchTool.Services;

public sealed class ExcelSearchHandler(ErrorLogger errorLogger) : ISearchHandler
{
    private readonly ErrorLogger _errorLogger = errorLogger;

    public async Task SearchAsync(FileInfo file, SearchContext context)
    {
        if (!CanContainMatch(file.FullName, context))
        {
            return;
        }

        try
        {
            await using var stream = File.Open(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var package = new ExcelPackage(stream);
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                context.CancellationToken.ThrowIfCancellationRequested();
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
                        if (string.IsNullOrEmpty(value) || !context.IsMatch(value))
                        {
                            continue;
                        }

                        context.Results.Add(new SearchResult
                        {
                            FileName = file.Name,
                            FilePath = file.FullName,
                            FileType = "Excel",
                            Location = $"{worksheet.Name}!R{row}C{column}",
                            Content = context.HighlightMatch(TrimForPreview(value))
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
            await _errorLogger.LogAsync($"Excel parse skipped: {file.FullName}", ex).ConfigureAwait(false);
        }
    }

    private static string TrimForPreview(string value)
    {
        return value.Length > 220 ? string.Concat(value.AsSpan(0, 220), "...") : value;
    }

    private static bool CanContainMatch(string filePath, SearchContext context)
    {
        if (context.Options.UseRegex)
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
                using var reader = new StreamReader(stream, Encoding.UTF8, true, 4096, false);
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (line is not null && line.Contains(context.Options.Query,
                            context.Options.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }

            return false;
        }
        catch
        {
            return true;
        }
    }
}
