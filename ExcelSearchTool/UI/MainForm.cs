using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using ExcelSearchTool.Models;
using ExcelSearchTool.Services;

namespace ExcelSearchTool.UI;

public sealed class MainForm : Form
{
    private const int SectionPadding = 10;
    private const int SectionSpacing = 12;
    private const int ControlHeight = 30;
    private const int ControlSpacing = 8;
    private const int FolderSectionHeight = 170;

    private readonly FolderManager _folderManager = new();
    private readonly HistoryService _historyService = new();
    private readonly SearchService _searchService;

    private readonly ComboBox _queryComboBox = new() { DropDownStyle = ComboBoxStyle.DropDown };
    private readonly ComboBox _modeComboBox = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly CheckBox _regexCheck = new() { Text = "Regex", AutoSize = true };
    private readonly CheckBox _caseSensitiveCheck = new() { Text = "Case", AutoSize = true };

    private readonly ListBox _foldersList = new();
    private readonly DataGridView _resultsGrid = new();
    private readonly Button _searchButton = new() { Text = "Search" };
    private readonly Button _stopButton = new() { Text = "Stop", Enabled = false };
    private readonly Button _exportButton = new() { Text = "Export" };
    private readonly Button _addFolderButton = new() { Text = "Add Folder" };
    private readonly Button _removeFolderButton = new() { Text = "Remove" };
    private readonly Button _clearFoldersButton = new() { Text = "Clear All" };
    private readonly ProgressBar _progressBar = new() { Style = ProgressBarStyle.Continuous, Minimum = 0, Maximum = 100 };
    private readonly Label _statusLabel = new() { Text = "Ready", AutoSize = true, TextAlign = ContentAlignment.MiddleRight };
    private readonly ToolTip _toolTip = new();

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
        MinimumSize = new Size(980, 680);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9.5F);
        BackColor = Color.FromArgb(245, 247, 250);

        BuildLayout();
        BindEvents();
    }

    protected override async void OnLoad(EventArgs e)
    {
        base.OnLoad(e);

        ConfigureModeCombo();
        var history = await _historyService.LoadAsync();
        _queryComboBox.Items.Clear();
        _queryComboBox.Items.AddRange(history.ToArray());
        SetCueBanner(_queryComboBox, "Type keyword, filename, or regex pattern...");
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(16),
            BackColor = BackColor
        };

        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, FolderSectionHeight));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var searchSection = BuildSearchSection();
        var folderSection = BuildFolderSection();
        var resultSection = BuildResultSection();
        var statusSection = BuildStatusSection();

        root.Controls.Add(searchSection, 0, 0);
        root.Controls.Add(folderSection, 0, 1);
        root.Controls.Add(resultSection, 0, 2);
        root.Controls.Add(statusSection, 0, 3);

        searchSection.Margin = new Padding(0, 0, 0, SectionSpacing);
        folderSection.Margin = new Padding(0, 0, 0, SectionSpacing);
        resultSection.Margin = new Padding(0, 0, 0, SectionSpacing);

        Controls.Add(root);
    }

    private Control BuildSearchSection()
    {
        var container = CreateSectionPanel();

        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 8,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };

        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));

        var keywordLabel = new Label
        {
            Text = "Keyword",
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 0, ControlSpacing, 0)
        };

        ConfigureStandardControl(_queryComboBox);
        _queryComboBox.Dock = DockStyle.Fill;
        _queryComboBox.Margin = new Padding(0, 0, ControlSpacing, 0);

        ConfigureStandardControl(_modeComboBox);
        _modeComboBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        _modeComboBox.Margin = new Padding(0, 0, ControlSpacing, 0);

        ConfigureCheck(_regexCheck);
        ConfigureCheck(_caseSensitiveCheck);

        ConfigureActionButton(_searchButton, Color.FromArgb(37, 99, 235), Color.White);
        ConfigureActionButton(_stopButton, Color.FromArgb(220, 38, 38), Color.White);
        ConfigureActionButton(_exportButton, Color.FromArgb(229, 231, 235), Color.FromArgb(31, 41, 55));

        grid.Controls.Add(keywordLabel, 0, 0);
        grid.Controls.Add(_queryComboBox, 1, 0);
        grid.Controls.Add(_modeComboBox, 2, 0);
        grid.Controls.Add(_regexCheck, 3, 0);
        grid.Controls.Add(_caseSensitiveCheck, 4, 0);
        grid.Controls.Add(_searchButton, 5, 0);
        grid.Controls.Add(_stopButton, 6, 0);
        grid.Controls.Add(_exportButton, 7, 0);

        container.Controls.Add(grid);

        _toolTip.SetToolTip(_queryComboBox, "Enter keyword, filename, or regex pattern");
        _toolTip.SetToolTip(_modeComboBox, "Choose whether to search content, filename, or both");
        _toolTip.SetToolTip(_regexCheck, "Enable regex matching");
        _toolTip.SetToolTip(_caseSensitiveCheck, "Match uppercase/lowercase exactly");
        _toolTip.SetToolTip(_exportButton, "Export current results to CSV");

        return container;
    }

    private Control BuildFolderSection()
    {
        var container = CreateSectionPanel();

        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };

        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));

        _foldersList.Dock = DockStyle.Fill;
        _foldersList.IntegralHeight = false;
        _foldersList.Margin = new Padding(0, 0, SectionSpacing, 0);

        var buttonStack = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            RowCount = 3,
            AutoSize = true,
            Padding = Padding.Empty,
            Margin = Padding.Empty
        };
        buttonStack.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        buttonStack.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        buttonStack.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        ConfigureSideButton(_addFolderButton);
        ConfigureSideButton(_removeFolderButton);
        ConfigureSideButton(_clearFoldersButton);

        buttonStack.Controls.Add(_addFolderButton, 0, 0);
        buttonStack.Controls.Add(_removeFolderButton, 0, 1);
        buttonStack.Controls.Add(_clearFoldersButton, 0, 2);

        _foldersList.DataSource = _folderManager.Folders;

        grid.Controls.Add(_foldersList, 0, 0);
        grid.Controls.Add(buttonStack, 1, 0);

        container.Controls.Add(grid);
        return container;
    }

    private Control BuildResultSection()
    {
        var container = CreateSectionPanel();
        ConfigureResultsGrid();
        container.Controls.Add(_resultsGrid);
        return container;
    }

    private Control BuildStatusSection()
    {
        var container = CreateSectionPanel();

        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };

        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        _progressBar.Dock = DockStyle.Fill;
        _progressBar.Margin = new Padding(0, 0, SectionSpacing, 0);
        _statusLabel.Anchor = AnchorStyles.Right;

        grid.Controls.Add(_progressBar, 0, 0);
        grid.Controls.Add(_statusLabel, 1, 0);

        container.Controls.Add(grid);
        return container;
    }

    private Panel CreateSectionPanel()
    {
        return new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(SectionPadding)
        };
    }

    private void ConfigureModeCombo()
    {
        if (_modeComboBox.Items.Count > 0)
        {
            return;
        }

        _modeComboBox.Items.AddRange(["Content + File Name", "Content Only", "File Name Only"]);
        _modeComboBox.SelectedIndex = 0;
    }

    private static void ConfigureStandardControl(Control control)
    {
        control.Height = ControlHeight;
    }

    private static void ConfigureCheck(CheckBox check)
    {
        check.Anchor = AnchorStyles.Left;
        check.Margin = new Padding(0, 0, ControlSpacing, 0);
    }

    private static void ConfigureActionButton(Button button, Color backColor, Color foreColor)
    {
        button.Height = ControlHeight;
        button.Dock = DockStyle.Fill;
        button.Margin = new Padding(0);
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 0;
        button.BackColor = backColor;
        button.ForeColor = foreColor;
        button.Cursor = Cursors.Hand;
    }

    private static void ConfigureSideButton(Button button)
    {
        button.Dock = DockStyle.Top;
        button.Height = ControlHeight;
        button.Margin = new Padding(0, 0, 0, ControlSpacing);
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
        button.BackColor = Color.White;
        button.Cursor = Cursors.Hand;
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
        _resultsGrid.AllowUserToResizeRows = false;
        _resultsGrid.RowHeadersVisible = false;
        _resultsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _resultsGrid.MultiSelect = false;
        _resultsGrid.DataSource = _results;
        _resultsGrid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254);
        _resultsGrid.DefaultCellStyle.SelectionForeColor = Color.Black;
        _resultsGrid.ColumnHeadersDefaultCellStyle.Font = new Font(Font, FontStyle.Bold);

        EnableDoubleBuffering(_resultsGrid);

        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.FileName), HeaderText = "File", Width = 180 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.MatchType), HeaderText = "Type", Width = 110 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.SheetName), HeaderText = "Sheet", Width = 140 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.Row), HeaderText = "Row", Width = 70 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.Column), HeaderText = "Col", Width = 70 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.MatchDetail), HeaderText = "Preview", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.FilePath), HeaderText = "Path", Width = 260 });
    }

    private static void EnableDoubleBuffering(DataGridView grid)
    {
        typeof(DataGridView).InvokeMember(
            "DoubleBuffered",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty,
            null,
            grid,
            [true]);
    }

    private void BindEvents()
    {
        _searchButton.Click += async (_, _) => await StartSearchAsync();
        _stopButton.Click += (_, _) => CancelActiveSearch();
        _exportButton.Click += (_, _) => ExportResults();

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
            if (!string.IsNullOrWhiteSpace(_queryComboBox.Text) && ContainsFocus)
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

        var total = Math.Max(1, CountExcelFiles(options.Folders));
        var progress = new Progress<SearchProgress>(p =>
        {
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

        var selectedMode = _modeComboBox.SelectedItem?.ToString() ?? "Content + File Name";
        var searchContent = selectedMode is "Content + File Name" or "Content Only";
        var searchFileName = selectedMode is "Content + File Name" or "File Name Only";

        return new SearchOptions
        {
            Query = query,
            SearchContent = searchContent,
            SearchFileName = searchFileName,
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

    private void ExportResults()
    {
        if (_results.Count == 0)
        {
            _statusLabel.Text = "No results to export.";
            return;
        }

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = $"ExcelSearchResults-{DateTime.Now:yyyyMMdd-HHmmss}.csv"
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        var csv = new StringBuilder();
        csv.AppendLine("FileName,MatchType,Sheet,Row,Column,MatchDetail,FilePath");
        foreach (var result in _results)
        {
            csv.AppendLine(string.Join(',',
                Csv(result.FileName),
                Csv(result.MatchType),
                Csv(result.SheetName ?? string.Empty),
                Csv(result.Row?.ToString() ?? string.Empty),
                Csv(result.Column?.ToString() ?? string.Empty),
                Csv(result.MatchDetail),
                Csv(result.FilePath)));
        }

        File.WriteAllText(dialog.FileName, csv.ToString(), Encoding.UTF8);
        _statusLabel.Text = $"Exported {_results.Count} rows.";
    }

    private static string Csv(string value)
    {
        if (value.Contains('"') || value.Contains(',') || value.Contains('\n'))
        {
            return $"\"{value.Replace("\"", "\"\"")}\"";
        }

        return value;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

    private static void SetCueBanner(ComboBox comboBox, string text)
    {
        const int emSetCueBanner = 0x1501;
        if (comboBox.IsHandleCreated)
        {
            _ = SendMessage(comboBox.Handle, emSetCueBanner, (IntPtr)1, text);
            return;
        }

        comboBox.HandleCreated += (_, _) => _ = SendMessage(comboBox.Handle, emSetCueBanner, (IntPtr)1, text);
    }
}
