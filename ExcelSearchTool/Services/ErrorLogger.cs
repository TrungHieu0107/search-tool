using System.Text;

namespace ExcelSearchTool.Services;

public sealed class ErrorLogger
{
    private readonly string _logFilePath;
    private readonly SemaphoreSlim _writeLock = new(1, 1);

    public ErrorLogger()
    {
        var basePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelSearchTool");
        Directory.CreateDirectory(basePath);
        _logFilePath = Path.Combine(basePath, "errors.log");
    }

    public async Task LogAsync(string message, Exception? ex = null)
    {
        var sb = new StringBuilder();
        sb.Append('[').Append(DateTimeOffset.Now).Append("] ").Append(message);
        if (ex is not null)
        {
            sb.AppendLine();
            sb.AppendLine(ex.ToString());
        }

        await _writeLock.WaitAsync().ConfigureAwait(false);
        try
        {
            await File.AppendAllTextAsync(_logFilePath, sb.AppendLine().AppendLine().ToString()).ConfigureAwait(false);
        }
        finally
        {
            _writeLock.Release();
        }
    }
}
