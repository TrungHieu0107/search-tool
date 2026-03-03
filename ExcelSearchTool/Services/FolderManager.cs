using System.ComponentModel;

namespace ExcelSearchTool.Services;

public sealed class FolderManager
{
    private readonly BindingList<string> _folders = [];

    public BindingList<string> Folders => _folders;

    public void Add(string folder)
    {
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            return;
        }

        if (_folders.Any(existing => string.Equals(existing, folder, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        _folders.Add(folder);
    }

    public void Remove(string folder)
    {
        var match = _folders.FirstOrDefault(existing => string.Equals(existing, folder, StringComparison.OrdinalIgnoreCase));
        if (match is not null)
        {
            _folders.Remove(match);
        }
    }

    public void Clear() => _folders.Clear();
}
