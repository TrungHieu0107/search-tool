using System.ComponentModel;
using System.Text.RegularExpressions;
using ExcelSearchTool.Models;
using ExcelSearchTool.Services;

namespace ExcelSearchTool.UI;

public sealed class MainForm : Form
{
    private readonly FolderManager _folderManager = new();
    private readonly HistoryService _historyService = new();
    private readonly SearchService _searchService;

    private readonly ComboBox _queryComboBox = new() { DropDownStyle = ComboBoxStyle.DropDown };
    private readonly CheckBox _contentSearchCheck = new() { Text = "Content Search", Checked = true, AutoSize = true };
    private readonly CheckBox _fileNameSearchCheck = new() { Text = "File Name Search", Checked = true, AutoSize = true };
    private readonly CheckBox _regexCheck = new() { Text = "Regex", AutoSize = true };
    private readonly CheckBox _caseSensitiveCheck = new() { Text = "Case Sensitive", AutoSize = true };

    private readonly ListBox _foldersList = new();
    private readonly DataGridView _resultsGrid = new();
    private readonly Button _searchButton = new() { Text = "Search" };
    private readonly Button _stopButton = new() { Text = "Stop", Enabled = false };
    private readonly Button _addFolderButton = new() { Text = "Add Folder" };
    private readonly Button _removeFolderButton = new() { Text = "Remove" };
    private readonly Button _clearFoldersButton = new() { Text = "Clear" };
    private readonly ProgressBar _progressBar = new() { Style = ProgressBarStyle.Continuous, Minimum = 0, Maximum = 100 };
    private readonly Label _statusLabel = new() { Text = "Ready", AutoSize = true };

    private readonly Timer _debounceTimer = new() { Interval = 400 };
    private readonly BindingList<SearchResult> _results = [];
    private CancellationTokenSource? _searchCts;

    public MainForm()
    {
        var errorLogger = new ErrorLogger();
        _searchService = new SearchService(new FileNameSearchService(), new ContentSearchService(errorLogger), errorLogger);

        Text = "Excel Search Tool";
        Width = 1200;
        Height = 800;
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10F);
        BackColor = Color.FromArgb(248, 250, 252);

        BuildLayout();
        BindEvents();
    }

    protected override async void OnLoad(EventArgs e)
    {
        base.OnLoad(e);

        var history = await _historyService.LoadAsync();
        _queryComboBox.Items.Clear();
        _queryComboBox.Items.AddRange(history.ToArray());
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(16)
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 130));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var searchPanel = new Panel { Dock = DockStyle.Top, Height = 95, Padding = new Padding(10), BackColor = Color.White };
        var queryLabel = new Label { Text = "Search Query", AutoSize = true, Location = new Point(8, 8) };
        _queryComboBox.SetBounds(8, 30, 520, 30);

        var checksFlow = new FlowLayoutPanel { Location = new Point(540, 30), AutoSize = true };
        checksFlow.Controls.AddRange([_contentSearchCheck, _fileNameSearchCheck, _regexCheck, _caseSensitiveCheck]);

        _searchButton.SetBounds(8, 62, 100, 30);
        _stopButton.SetBounds(120, 62, 100, 30);

        searchPanel.Controls.AddRange([queryLabel, _queryComboBox, checksFlow, _searchButton, _stopButton]);

        var folderPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(10) };
        var foldersTitle = new Label { Text = "Search Folders", AutoSize = true, Location = new Point(8, 8) };
        _foldersList.SetBounds(8, 30, 1040, 80);
        _foldersList.DataSource = _folderManager.Folders;
        _addFolderButton.SetBounds(1060, 30, 110, 26);
        _removeFolderButton.SetBounds(1060, 60, 110, 26);
        _clearFoldersButton.SetBounds(1060, 90, 110, 26);
        folderPanel.Controls.AddRange([foldersTitle, _foldersList, _addFolderButton, _removeFolderButton, _clearFoldersButton]);

        ConfigureResultsGrid();

        var statusPanel = new Panel { Dock = DockStyle.Fill, Height = 40, BackColor = Color.White, Padding = new Padding(8) };
        _progressBar.SetBounds(8, 8, 350, 22);
        _statusLabel.SetBounds(370, 11, 700, 20);
        statusPanel.Controls.AddRange([_progressBar, _statusLabel]);

        root.Controls.Add(searchPanel, 0, 0);
        root.Controls.Add(folderPanel, 0, 1);
        root.Controls.Add(_resultsGrid, 0, 2);
        root.Controls.Add(statusPanel, 0, 3);

        Controls.Add(root);
    }

    private void ConfigureResultsGrid()
    {
        _resultsGrid.Dock = DockStyle.Fill;
        _resultsGrid.BackgroundColor = Color.White;
        _resultsGrid.BorderStyle = BorderStyle.None;
        _resultsGrid.AutoGenerateColumns = false;
        _resultsGrid.ReadOnly = true;
        _resultsGrid.AllowUserToAddRows = false;
        _resultsGrid.AllowUserToDeleteRows = false;
        _resultsGrid.RowHeadersVisible = false;
        _resultsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _resultsGrid.DataSource = _results;

        EnableDoubleBuffering(_resultsGrid);

        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.FileName), HeaderText = "File", Width = 220 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.MatchType), HeaderText = "Type", Width = 110 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.SheetName), HeaderText = "Sheet", Width = 140 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.Row), HeaderText = "Row", Width = 80 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.Column), HeaderText = "Col", Width = 80 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.MatchDetail), HeaderText = "Preview", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.FilePath), HeaderText = "Path", Width = 250 });
    }

    private static void EnableDoubleBuffering(DataGridView grid)
    {
        typeof(DataGridView).InvokeMember("DoubleBuffered",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty,
            null,
            grid,
            [true]);
    }

    private void BindEvents()
    {
        _searchButton.Click += async (_, _) => await StartSearchAsync();
        _stopButton.Click += (_, _) => CancelActiveSearch();
        _addFolderButton.Click += (_, _) => AddFolder();
        _removeFolderButton.Click += (_, _) => RemoveSelectedFolder();
        _clearFoldersButton.Click += (_, _) => _folderManager.Clear();

        _queryComboBox.TextChanged += (_, _) =>
        {
            _debounceTimer.Stop();
            _debounceTimer.Start();
        };

        _debounceTimer.Tick += async (_, _) =>
        {
            _debounceTimer.Stop();
            if (Focused && !string.IsNullOrWhiteSpace(_queryComboBox.Text))
            {
                await StartSearchAsync();
            }
        };
    }

    private void AddFolder()
    {
        using var dialog = new FolderBrowserDialog();
        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _folderManager.Add(dialog.SelectedPath);
        }
    }

    private void RemoveSelectedFolder()
    {
        if (_foldersList.SelectedItem is string folder)
        {
            _folderManager.Remove(folder);
        }
    }

    private async Task StartSearchAsync()
    {
        CancelActiveSearch();

        var options = BuildSearchOptions();
        if (options is null)
        {
            return;
        }

        _searchCts = new CancellationTokenSource();
        _searchButton.Enabled = false;
        _stopButton.Enabled = true;
        _results.Clear();
        _progressBar.Value = 0;

        var progress = new Progress<SearchProgress>(p =>
        {
            var total = Math.Max(1, CountExcelFiles(options.Folders));
            _progressBar.Value = Math.Min(100, (int)Math.Round((double)p.FilesScanned / total * 100));
            _statusLabel.Text = $"Scanned: {p.FilesScanned}/{total} | Matches: {p.Matches}";
        });

        try
        {
            await _historyService.SaveAsync(options.Query);
            var results = await _searchService.SearchAsync(options, progress, _searchCts.Token);
            foreach (var result in results)
            {
                _results.Add(result);
            }

            _statusLabel.Text = $"Done. {results.Count} matches.";
        }
        catch (OperationCanceledException)
        {
            _statusLabel.Text = "Search canceled.";
        }
        finally
        {
            _searchButton.Enabled = true;
            _stopButton.Enabled = false;
            _searchCts?.Dispose();
            _searchCts = null;
        }
    }

    private SearchOptions? BuildSearchOptions()
    {
        var query = _queryComboBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(query))
        {
            _statusLabel.Text = "Enter a query.";
            return null;
        }

        if (_folderManager.Folders.Count == 0)
        {
            _statusLabel.Text = "Add at least one folder.";
            return null;
        }

        if (_regexCheck.Checked)
        {
            try
            {
                _ = new Regex(query);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Invalid regex: {ex.Message}", "Regex Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
        }

        return new SearchOptions
        {
            Query = query,
            SearchContent = _contentSearchCheck.Checked,
            SearchFileName = _fileNameSearchCheck.Checked,
            UseRegex = _regexCheck.Checked,
            CaseSensitive = _caseSensitiveCheck.Checked,
            Folders = _folderManager.Folders.ToArray()
        };
    }

    private static int CountExcelFiles(IEnumerable<string> folders)
    {
        return folders.Sum(folder => Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories)
            .Count(path => path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                || path.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)
                || path.EndsWith(".xlsb", StringComparison.OrdinalIgnoreCase)));
    }

    private void CancelActiveSearch()
    {
        if (_searchCts is { IsCancellationRequested: false })
        {
            _searchCts.Cancel();
        }
    }
}
