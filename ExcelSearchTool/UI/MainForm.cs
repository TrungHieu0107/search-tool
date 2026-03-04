using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using ExcelSearchTool.Models;
using ExcelSearchTool.Services;

namespace ExcelSearchTool.UI;

public sealed class MainForm : Form
{
    private const int SectionPadding = 12;
    private const int SectionSpacing = 12;
    private const int ControlHeight = 36; // 32-48px target for clickability and touch comfort.
    private const int SmallControlHeight = 30;
    private const int ControlSpacing = 8;
    private const int FolderSectionHeight = 168;
    private const int StatusSectionHeight = 26; // Keep processing/status bar slim (20-30px).

    private readonly FolderManager _folderManager = new();
    private readonly HistoryService _historyService = new();
    private readonly SearchService _searchService;

    private readonly ComboBox _queryComboBox = new() { DropDownStyle = ComboBoxStyle.DropDown };
    private readonly ComboBox _modeComboBox = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly CheckBox _regexCheck = new() { Text = "Regex", AutoSize = true };
    private readonly CheckBox _caseSensitiveCheck = new() { Text = "Case", AutoSize = true };
    private readonly CheckBox _excelFilterCheck = new() { Text = "Excel", Checked = true, AutoSize = true };
    private readonly CheckBox _textFilterCheck = new() { Text = "Text", Checked = true, AutoSize = true };
    private readonly CheckBox _codeFilterCheck = new() { Text = "Code", Checked = true, AutoSize = true };
    private readonly CheckBox _darkModeCheck = new() { Text = "Dark", AutoSize = true };

    private readonly ListBox _foldersList = new();
    private readonly DataGridView _resultsGrid = new();
    private readonly Button _searchButton = new() { Text = "Search" };
    private readonly Button _stopButton = new() { Text = "Stop", Enabled = false };
    private readonly Button _exportButton = new() { Text = "Export" };
    private readonly Button _addFolderButton = new() { Text = "Add Folder" };
    private readonly Button _removeFolderButton = new() { Text = "Remove" };
    private readonly Button _clearFoldersButton = new() { Text = "Clear All" };
    private readonly ProgressBar _progressBar = new() { Style = ProgressBarStyle.Continuous, Minimum = 0, Maximum = 100, Visible = false };
    private readonly Label _statusLabel = new() { Text = "Ready", AutoSize = true, TextAlign = ContentAlignment.MiddleLeft };
    private readonly ToolTip _toolTip = new() { AutoPopDelay = 6000, InitialDelay = 300, ReshowDelay = 150 };

    private readonly System.Windows.Forms.Timer _debounceTimer = new() { Interval = 400 };
    private readonly BindingList<SearchResult> _results = new();
    private readonly Label _resultsCountLabel = new() { Text = "0 matches", AutoSize = true };

    private readonly Color _lightSurface = Color.FromArgb(255, 255, 255);
    private readonly Color _lightPanel = Color.FromArgb(245, 247, 250);
    private readonly Color _lightText = Color.FromArgb(17, 24, 39);
    private readonly Color _lightSubtle = Color.FromArgb(100, 116, 139);

    private readonly Color _darkSurface = Color.FromArgb(31, 41, 55);
    private readonly Color _darkPanel = Color.FromArgb(17, 24, 39);
    private readonly Color _darkText = Color.FromArgb(226, 232, 240);
    private readonly Color _darkSubtle = Color.FromArgb(148, 163, 184);

    private CancellationTokenSource? _searchCts;

    public MainForm()
    {
        var errorLogger = new ErrorLogger();
        _searchService = new SearchService(errorLogger);

        Text = "Universal File Search Tool";
        Width = 1300;
        Height = 820;
        MinimumSize = new Size(1040, 700);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9.5F);
        KeyPreview = true;

        BuildLayout();
        BindEvents();
        ApplyTheme(isDark: false);
    }

    protected override async void OnLoad(EventArgs e)
    {
        base.OnLoad(e);

        ConfigureModeCombo();
        var history = await _historyService.LoadAsync();
        _queryComboBox.Items.Clear();
        _queryComboBox.Items.AddRange(history.ToArray());
        SetCueBanner(_queryComboBox, "Search in Excel, text, and source code files...");
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
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, FolderSectionHeight));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, StatusSectionHeight));

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
            ColumnCount = 12,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };

        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 84));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 84));

        var keywordLabel = new Label { Text = "Keyword", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, ControlSpacing, 0) };
        ConfigureStandardControl(_queryComboBox);
        _queryComboBox.Dock = DockStyle.Fill;
        _queryComboBox.Margin = new Padding(0, 0, ControlSpacing, 0);

        ConfigureStandardControl(_modeComboBox);
        _modeComboBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        _modeComboBox.Margin = new Padding(0, 0, ControlSpacing, 0);

        ConfigureCheck(_regexCheck);
        ConfigureCheck(_caseSensitiveCheck);
        ConfigureCheck(_excelFilterCheck);
        ConfigureCheck(_textFilterCheck);
        ConfigureCheck(_codeFilterCheck);
        ConfigureCheck(_darkModeCheck);

        ConfigureActionButton(_searchButton, Color.FromArgb(30, 64, 175), Color.White);
        ConfigureActionButton(_stopButton, Color.FromArgb(185, 28, 28), Color.White);
        ConfigureActionButton(_exportButton, Color.FromArgb(226, 232, 240), Color.FromArgb(30, 41, 59));

        grid.Controls.Add(keywordLabel, 0, 0);
        grid.Controls.Add(_queryComboBox, 1, 0);
        grid.Controls.Add(_modeComboBox, 2, 0);
        grid.Controls.Add(_regexCheck, 3, 0);
        grid.Controls.Add(_caseSensitiveCheck, 4, 0);
        grid.Controls.Add(_excelFilterCheck, 5, 0);
        grid.Controls.Add(_textFilterCheck, 6, 0);
        grid.Controls.Add(_codeFilterCheck, 7, 0);
        grid.Controls.Add(_darkModeCheck, 8, 0);
        grid.Controls.Add(_searchButton, 9, 0);
        grid.Controls.Add(_stopButton, 10, 0);
        grid.Controls.Add(_exportButton, 11, 0);

        container.Controls.Add(grid);

        // Tooltips and keyboard hints improve discoverability.
        _toolTip.SetToolTip(_queryComboBox, "Enter keyword, filename, or regex pattern (Enter/Ctrl+F = Search)");
        _toolTip.SetToolTip(_modeComboBox, "Search content and/or file names");
        _toolTip.SetToolTip(_excelFilterCheck, "Include .xlsx files");
        _toolTip.SetToolTip(_textFilterCheck, "Include .txt, .log, .csv files");
        _toolTip.SetToolTip(_codeFilterCheck, "Include source/code-like file types");
        _toolTip.SetToolTip(_searchButton, "Run search (Enter or Ctrl+F)");
        _toolTip.SetToolTip(_stopButton, "Stop active search (Esc)");
        _toolTip.SetToolTip(_exportButton, "Export current results as CSV (Ctrl+E)");
        _toolTip.SetToolTip(_darkModeCheck, "Toggle dark theme");

        return container;
    }

    private Control BuildFolderSection()
    {
        var container = CreateSectionPanel();
        var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty, Padding = Padding.Empty };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));

        _foldersList.Dock = DockStyle.Fill;
        _foldersList.IntegralHeight = false;
        _foldersList.Margin = new Padding(0, 0, SectionSpacing, 0);

        var buttonStack = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 1, RowCount = 3, AutoSize = true, Padding = Padding.Empty, Margin = Padding.Empty };
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

        _toolTip.SetToolTip(_addFolderButton, "Add a root folder (Ctrl+O)");
        _toolTip.SetToolTip(_removeFolderButton, "Remove selected folder (Del)");
        _toolTip.SetToolTip(_clearFoldersButton, "Remove all configured folders");

        return container;
    }

    private Control BuildResultSection()
    {
        var container = CreateSectionPanel();
        var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Margin = Padding.Empty, Padding = Padding.Empty };
        grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var header = new Label
        {
            Text = "Search results",
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold),
            Dock = DockStyle.Top,
            Margin = new Padding(0, 0, 0, 8)
        };

        _resultsCountLabel.Dock = DockStyle.Top;
        _resultsCountLabel.Margin = new Padding(0, 0, 0, 8);
        _resultsCountLabel.ForeColor = _lightSubtle;

        var headerPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            AutoSize = true,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };

        headerPanel.Controls.Add(header);
        headerPanel.Controls.Add(new Label { Text = "  •  ", ForeColor = _lightSubtle, AutoSize = true, Margin = new Padding(0, 2, 0, 0) });
        headerPanel.Controls.Add(_resultsCountLabel);

        ConfigureResultsGrid();

        grid.Controls.Add(headerPanel, 0, 0);
        grid.Controls.Add(_resultsGrid, 0, 1);

        container.Controls.Add(grid);
        return container;
    }

    private Control BuildStatusSection()
    {
        var container = CreateSectionPanel();
        container.Padding = new Padding(8, 4, 8, 4);

        var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty, Padding = Padding.Empty };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        _progressBar.Dock = DockStyle.Fill;
        _progressBar.Height = SmallControlHeight - 8;
        _progressBar.Margin = new Padding(0, 0, SectionSpacing, 0);
        _statusLabel.Anchor = AnchorStyles.Right;

        grid.Controls.Add(_progressBar, 0, 0);
        grid.Controls.Add(_statusLabel, 1, 0);
        container.Controls.Add(grid);
        return container;
    }

    private Panel CreateSectionPanel() => new()
    {
        Dock = DockStyle.Fill,
        Padding = new Padding(SectionPadding)
    };

    private void ConfigureModeCombo()
    {
        if (_modeComboBox.Items.Count > 0) return;
        _modeComboBox.Items.AddRange(["Content + File Name", "Content Only", "File Name Only"]);
        _modeComboBox.SelectedIndex = 0;
    }

    private static void ConfigureStandardControl(Control control) => control.Height = ControlHeight;

    private static void ConfigureCheck(CheckBox check)
    {
        check.Anchor = AnchorStyles.Left;
        check.Margin = new Padding(0, 0, ControlSpacing, 0);
    }

    private static void ConfigureActionButton(Button button, Color backColor, Color foreColor)
    {
        button.Height = ControlHeight;
        button.Dock = DockStyle.Fill;
        button.Margin = Padding.Empty;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 0;
        button.BackColor = backColor;
        button.ForeColor = foreColor;
        button.Cursor = Cursors.Hand;
    }

    private static void ConfigureSideButton(Button button)
    {
        button.Dock = DockStyle.Top;
        button.Height = SmallControlHeight;
        button.Margin = new Padding(0, 0, 0, ControlSpacing);
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
        button.BackColor = Color.White;
        button.Cursor = Cursors.Hand;
    }

    private void ConfigureResultsGrid()
    {
        _resultsGrid.Dock = DockStyle.Fill;
        _resultsGrid.BorderStyle = BorderStyle.None;
        _resultsGrid.AutoGenerateColumns = false;
        _resultsGrid.ReadOnly = true;
        _resultsGrid.AllowUserToAddRows = false;
        _resultsGrid.AllowUserToDeleteRows = false;
        _resultsGrid.AllowUserToResizeRows = false;
        _resultsGrid.AllowUserToResizeColumns = true; // Explicitly support user-resizable columns.
        _resultsGrid.RowHeadersVisible = false;
        _resultsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _resultsGrid.MultiSelect = false;
        _resultsGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
        _resultsGrid.RowTemplate.Height = 34;
        _resultsGrid.ColumnHeadersHeight = 36;
        _resultsGrid.DataSource = _results;
        _resultsGrid.DefaultCellStyle.Padding = new Padding(6, 4, 6, 4);
        _resultsGrid.ColumnHeadersDefaultCellStyle.Padding = new Padding(6, 6, 6, 6);

        EnableDoubleBuffering(_resultsGrid);

        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.FileName), HeaderText = "File Name", Width = 180, MinimumWidth = 120 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.FilePath), HeaderText = "Full Path", Width = 320, MinimumWidth = 200 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.FileType), HeaderText = "Type", Width = 90, MinimumWidth = 70 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.Location), HeaderText = "Location", Width = 150, MinimumWidth = 110 });
        _resultsGrid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(SearchResult.Content), HeaderText = "Content", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 260 });
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
        _exportButton.Click += (_, _) => ExportResults();

        _addFolderButton.Click += (_, _) => AddFolder();
        _removeFolderButton.Click += (_, _) => RemoveSelectedFolder();
        _clearFoldersButton.Click += (_, _) => _folderManager.Clear();
        _darkModeCheck.CheckedChanged += (_, _) => ApplyTheme(_darkModeCheck.Checked);

        _queryComboBox.TextChanged += (_, _) =>
        {
            _debounceTimer.Stop();
            _debounceTimer.Start();
        };

        _resultsGrid.CellDoubleClick += (_, e) => OpenSelectedResult(e.RowIndex);
        _results.ListChanged += (_, _) => _resultsCountLabel.Text = $"{_results.Count} matches";

        _debounceTimer.Tick += async (_, _) =>
        {
            _debounceTimer.Stop();
            if (!string.IsNullOrWhiteSpace(_queryComboBox.Text) && ContainsFocus)
            {
                await StartSearchAsync();
            }
        };

        KeyDown += async (_, e) =>
        {
            if ((e.Control && e.KeyCode == Keys.F) || e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                await StartSearchAsync();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                CancelActiveSearch();
            }
            else if (e.Control && e.KeyCode == Keys.E)
            {
                ExportResults();
            }
            else if (e.Control && e.KeyCode == Keys.O)
            {
                AddFolder();
            }
            else if (e.KeyCode == Keys.Delete && _foldersList.Focused)
            {
                RemoveSelectedFolder();
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
        _progressBar.Visible = true;

        var total = Math.Max(1, _searchService.CountEligibleFiles(options));
        var progress = new Progress<SearchProgress>(p =>
        {
            _progressBar.Value = Math.Min(100, (int)Math.Round((double)p.FilesScanned / total * 100));
            _statusLabel.Text = $"Scanning {p.FilesScanned}/{total} • {p.Matches} matches";
        });

        try
        {
            await _historyService.SaveAsync(options.Query);
            var results = await _searchService.SearchAsync(options, progress, _searchCts.Token);
            foreach (var result in results)
            {
                _results.Add(result);
            }

            _statusLabel.Text = $"Done, {results.Count} matches.";
        }
        catch (OperationCanceledException)
        {
            _statusLabel.Text = "Search canceled.";
        }
        catch (Exception ex)
        {
            _statusLabel.Text = "Search failed. Check logs for details.";
            MessageBox.Show(this, ex.Message, "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _searchButton.Enabled = true;
            _stopButton.Enabled = false;
            _progressBar.Visible = false; // Keep bottom bar subtle: hide when idle.
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

        if (!_excelFilterCheck.Checked && !_textFilterCheck.Checked && !_codeFilterCheck.Checked)
        {
            _statusLabel.Text = "Select at least one file type filter.";
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
        var searchFileName = selectedMode is "Content + File Name" or "File Name Only";

        return new SearchOptions
        {
            Query = query,
            SearchFileName = searchFileName,
            UseRegex = _regexCheck.Checked,
            CaseSensitive = _caseSensitiveCheck.Checked,
            IncludeExcel = _excelFilterCheck.Checked,
            IncludeText = _textFilterCheck.Checked,
            IncludeCode = _codeFilterCheck.Checked,
            Folders = _folderManager.Folders.ToArray()
        };
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

        using var dialog = new SaveFileDialog { Filter = "CSV Files (*.csv)|*.csv", FileName = $"UniversalSearchResults-{DateTime.Now:yyyyMMdd-HHmmss}.csv" };
        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        var csv = new StringBuilder();
        csv.AppendLine("FileName,FullPath,Type,Location,Content");
        foreach (var result in _results)
        {
            csv.AppendLine(string.Join(',', Csv(result.FileName), Csv(result.FilePath), Csv(result.FileType), Csv(result.Location), Csv(result.Content)));
        }

        File.WriteAllText(dialog.FileName, csv.ToString(), Encoding.UTF8);
        _statusLabel.Text = $"Exported {_results.Count} rows.";
    }

    private void OpenSelectedResult(int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= _results.Count)
        {
            return;
        }

        var item = _results[rowIndex];
        if (!File.Exists(item.FilePath))
        {
            _statusLabel.Text = "File no longer exists.";
            return;
        }

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = item.FilePath,
                UseShellExecute = true
            };
            Process.Start(psi);
        }
        catch (Exception ex)
        {
            _statusLabel.Text = $"Open failed: {ex.Message}";
        }
    }

    private void ApplyTheme(bool isDark)
    {
        var surface = isDark ? _darkSurface : _lightSurface;
        var panel = isDark ? _darkPanel : _lightPanel;
        var text = isDark ? _darkText : _lightText;
        var subtle = isDark ? _darkSubtle : _lightSubtle;

        BackColor = panel;
        ForeColor = text;

        foreach (var section in Controls.OfType<TableLayoutPanel>().SelectMany(r => r.Controls.OfType<Panel>()))
        {
            section.BackColor = surface;
            section.ForeColor = text;
        }

        _foldersList.BackColor = surface;
        _foldersList.ForeColor = text;
        _statusLabel.ForeColor = subtle;
        _resultsCountLabel.ForeColor = subtle;

        _resultsGrid.BackgroundColor = surface;
        _resultsGrid.GridColor = isDark ? Color.FromArgb(51, 65, 85) : Color.FromArgb(226, 232, 240);
        _resultsGrid.DefaultCellStyle.BackColor = surface;
        _resultsGrid.DefaultCellStyle.ForeColor = text;
        _resultsGrid.DefaultCellStyle.SelectionBackColor = isDark ? Color.FromArgb(30, 58, 138) : Color.FromArgb(219, 234, 254);
        _resultsGrid.DefaultCellStyle.SelectionForeColor = isDark ? Color.White : Color.Black;
        _resultsGrid.ColumnHeadersDefaultCellStyle.BackColor = isDark ? Color.FromArgb(30, 41, 59) : Color.FromArgb(248, 250, 252);
        _resultsGrid.ColumnHeadersDefaultCellStyle.ForeColor = text;
        _resultsGrid.EnableHeadersVisualStyles = false;

        foreach (var button in Controls.OfType<TableLayoutPanel>().SelectMany(x => x.Controls.OfType<Panel>()).SelectMany(x => x.Controls.OfType<TableLayoutPanel>()).SelectMany(x => x.Controls.OfType<Button>()))
        {
            if (button == _searchButton || button == _stopButton)
            {
                continue;
            }

            button.BackColor = isDark ? Color.FromArgb(51, 65, 85) : Color.FromArgb(226, 232, 240);
            button.ForeColor = text;
        }
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
