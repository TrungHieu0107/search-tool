using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using ExcelSearchTool.Models;

namespace ExcelSearchTool.Services;

public sealed class SearchProgress
{
    public int FilesScanned { get; init; }
    public int Matches { get; init; }
}

public sealed class SearchService(
    FileNameSearchService fileNameSearchService,
    ContentSearchService contentSearchService,
    ErrorLogger errorLogger)
{
    private static readonly string[] ExcelExtensions = [".xlsx", ".xlsm", ".xlsb"];
    private readonly FileNameSearchService _fileNameSearchService = fileNameSearchService;
    private readonly ContentSearchService _contentSearchService = contentSearchService;
    private readonly ErrorLogger _errorLogger = errorLogger;

    public async Task<IReadOnlyCollection<SearchResult>> SearchAsync(
        SearchOptions options,
        IProgress<SearchProgress>? progress,
        CancellationToken cancellationToken)
    {
        var files = options.Folders
            .SelectMany(folder => Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories))
            .Where(path => ExcelExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var results = new ConcurrentBag<SearchResult>();
        var regex = BuildRegex(options);
        var scanned = 0;

        await Parallel.ForEachAsync(files, cancellationToken, async (file, ct) =>
        {
            try
            {
                ct.ThrowIfCancellationRequested();

                if (_fileNameSearchService.IsMatch(file, options, regex))
                {
                    results.Add(_fileNameSearchService.ToResult(file));
                }

                if (options.SearchContent && _contentSearchService.CanContainMatch(file, options))
                {
                    var contentMatches = await _contentSearchService.FindMatchesAsync(file, options, regex, ct).ConfigureAwait(false);
                    foreach (var match in contentMatches)
                    {
                        results.Add(match);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                await _errorLogger.LogAsync($"Unhandled file search error: {file}", ex).ConfigureAwait(false);
            }
            finally
            {
                var currentScanned = Interlocked.Increment(ref scanned);
                progress?.Report(new SearchProgress
                {
                    FilesScanned = currentScanned,
                    Matches = results.Count
                });
            }
        }).ConfigureAwait(false);

        return results
            .OrderBy(item => item.FileName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(item => item.SheetIndex)
            .ThenBy(item => item.Row)
            .ToArray();
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
