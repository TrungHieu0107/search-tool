using System.Text;
using ExcelSearchTool.Models;

namespace ExcelSearchTool.Services;

public sealed class CodeSearchHandler(ErrorLogger errorLogger) : ISearchHandler
{
    private readonly ErrorLogger _errorLogger = errorLogger;

    public async Task SearchAsync(FileInfo file, SearchContext context)
    {
        if (IsLikelyBinary(file.FullName))
        {
            return;
        }

        try
        {
            await SearchWithEncodingAsync(file, context, new UTF8Encoding(true, true)).ConfigureAwait(false);
        }
        catch (DecoderFallbackException)
        {
            try
            {
                await SearchWithEncodingAsync(file, context, Encoding.Default).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                await _errorLogger.LogAsync($"Code file skipped due to encoding/read error: {file.FullName}", ex).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            await _errorLogger.LogAsync($"Code file skipped due to read error: {file.FullName}", ex).ConfigureAwait(false);
        }
    }

    private static async Task SearchWithEncodingAsync(FileInfo file, SearchContext context, Encoding encoding)
    {
        await using var stream = File.Open(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = new StreamReader(stream, encoding, true);

        var lineNo = 0;
        while (!reader.EndOfStream)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            var line = await reader.ReadLineAsync(context.CancellationToken).ConfigureAwait(false);
            lineNo++;

            if (line is null || !context.IsMatch(line))
            {
                continue;
            }

            context.Results.Add(new SearchResult
            {
                FileName = file.Name,
                FilePath = file.FullName,
                FileType = "Code",
                Location = $"Line {lineNo}",
                Content = context.HighlightMatch(TrimForPreview(line))
            });
        }
    }

    private static string TrimForPreview(string value)
    {
        return value.Length > 220 ? string.Concat(value.AsSpan(0, 220), "...") : value;
    }

    private static bool IsLikelyBinary(string path)
    {
        var buffer = new byte[1024];
        using var stream = File.OpenRead(path);
        var read = stream.Read(buffer, 0, buffer.Length);
        for (var i = 0; i < read; i++)
        {
            if (buffer[i] == 0)
            {
                return true;
            }
        }

        return false;
    }
}
