using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using ExcelSearchTool.Models;

namespace ExcelSearchTool.Services;

public sealed class SearchProgress
{
    public int FilesScanned { get; init; }
    public int Matches { get; init; }
}

public sealed class SearchService(ErrorLogger errorLogger)
{
    private static readonly HashSet<string> ExcelExtensions = [".xlsx"];
    private static readonly HashSet<string> TextExtensions = [".txt", ".log", ".csv"];
    private static readonly HashSet<string> CodeExtensions = [".cs", ".java", ".js", ".ts", ".json", ".xml", ".html", ".css"];

    private readonly ErrorLogger _errorLogger = errorLogger;
    private readonly ISearchHandler _excelHandler = new ExcelSearchHandler(errorLogger);
    private readonly ISearchHandler _textHandler = new TextSearchHandler(errorLogger);
    private readonly ISearchHandler _codeHandler = new CodeSearchHandler(errorLogger);

    public async Task<IReadOnlyCollection<SearchResult>> SearchAsync(
        SearchOptions options,
        IProgress<SearchProgress>? progress,
        CancellationToken cancellationToken)
    {
        var files = options.Folders
            .SelectMany(SafeEnumerateFiles)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(path => new FileInfo(path))
            .Where(file => IsSupported(file.Extension, options))
            .Where(file => file.Exists && file.Length <= options.MaxFileSizeBytes)
            .ToArray();

        var results = new ConcurrentBag<SearchResult>();
        var regex = BuildRegex(options);
        var scanned = 0;

        await Parallel.ForEachAsync(files, cancellationToken, async (file, ct) =>
        {
            try
            {
                ct.ThrowIfCancellationRequested();

                if (options.SearchFileName && IsNameMatch(file.Name, options, regex))
                {
                    results.Add(new SearchResult
                    {
                        FileName = file.Name,
                        FilePath = file.FullName,
                        FileType = ResolveTypeLabel(file.Extension),
                        Location = "File Name",
                        Content = file.Name
                    });
                }

                var context = new SearchContext
                {
                    Options = options,
                    Regex = regex,
                    CancellationToken = ct,
                    Results = results
                };

                await ResolveHandler(file.Extension).SearchAsync(file, context).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                await _errorLogger.LogAsync($"Unhandled search error: {file.FullName}", ex).ConfigureAwait(false);
            }
            finally
            {
                var current = Interlocked.Increment(ref scanned);
                progress?.Report(new SearchProgress { FilesScanned = current, Matches = results.Count });
            }
        }).ConfigureAwait(false);

        return results
            .OrderBy(r => r.FileName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.Location, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    public int CountEligibleFiles(SearchOptions options)
    {
        return options.Folders
            .SelectMany(SafeEnumerateFiles)
            .Select(path => new FileInfo(path))
            .Count(file => file.Exists && file.Length <= options.MaxFileSizeBytes && IsSupported(file.Extension, options));
    }

    private static IEnumerable<string> SafeEnumerateFiles(string folder)
    {
        try
        {
            return Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories);
        }
        catch
        {
            return [];
        }
    }

    private static bool IsSupported(string extension, SearchOptions options)
    {
        extension = extension.ToLowerInvariant();
        return (options.IncludeExcel && ExcelExtensions.Contains(extension))
            || (options.IncludeText && TextExtensions.Contains(extension))
            || (options.IncludeCode && CodeExtensions.Contains(extension));
    }

    private ISearchHandler ResolveHandler(string extension)
    {
        extension = extension.ToLowerInvariant();
        if (ExcelExtensions.Contains(extension)) return _excelHandler;
        if (TextExtensions.Contains(extension)) return _textHandler;
        if (CodeExtensions.Contains(extension)) return _codeHandler;
        return _textHandler;
    }

    private static string ResolveTypeLabel(string extension)
    {
        extension = extension.ToLowerInvariant();
        if (ExcelExtensions.Contains(extension)) return "Excel";
        if (TextExtensions.Contains(extension)) return "Text";
        if (CodeExtensions.Contains(extension)) return "Code";
        return "Unknown";
    }

    private static bool IsNameMatch(string fileName, SearchOptions options, Regex? regex)
    {
        return regex is not null
            ? regex.IsMatch(fileName)
            : fileName.Contains(options.Query, options.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
    }

    private static Regex? BuildRegex(SearchOptions options)
    {
        if (!options.UseRegex)
        {
            return null;
        }

        var regexOptions = RegexOptions.Compiled;
        if (!options.CaseSensitive)
        {
            regexOptions |= RegexOptions.IgnoreCase;
        }

        return new Regex(options.Query, regexOptions, TimeSpan.FromSeconds(2));
    }
}
