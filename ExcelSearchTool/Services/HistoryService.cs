using System.Text.Json;

namespace ExcelSearchTool.Services;

public sealed class HistoryService
{
    private const int MaxEntries = 20;
    private readonly string _filePath;

    public HistoryService()
    {
        var appDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelSearchTool");
        Directory.CreateDirectory(appDir);
        _filePath = Path.Combine(appDir, "search-history.json");
    }

    public async Task<IReadOnlyList<string>> LoadAsync()
    {
        if (!File.Exists(_filePath))
        {
            return [];
        }

        await using var stream = File.OpenRead(_filePath);
        var data = await JsonSerializer.DeserializeAsync<List<string>>(stream).ConfigureAwait(false);
        return data?.Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().ToList() ?? [];
    }

    public async Task SaveAsync(string query)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            return;
        }

        var history = (await LoadAsync().ConfigureAwait(false)).ToList();
        history.RemoveAll(value => string.Equals(value, query, StringComparison.OrdinalIgnoreCase));
        history.Insert(0, query);
        if (history.Count > MaxEntries)
        {
            history = history.Take(MaxEntries).ToList();
        }

        await File.WriteAllTextAsync(_filePath, JsonSerializer.Serialize(history)).ConfigureAwait(false);
    }
}
