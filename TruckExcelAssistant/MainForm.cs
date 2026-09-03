using TruckExcelAssistant.Services;

namespace TruckExcelAssistant;

public sealed class MainForm : Form
{
    private readonly Panel _contentHost = new();
    private readonly Label _pageTitle = new();
    private readonly Label _databaseStatus = new();
    private readonly Dictionary<string, Button> _navigationButtons = new(StringComparer.OrdinalIgnoreCase);
    private readonly DatabaseService _database;
    private readonly NewHaulControl _newHaulControl;
    private readonly HaulListControl _haulListControl;
    private readonly ExcelOutputControl _invoiceControl;
    private readonly ExcelOutputControl _ledgerControl;
    private readonly InvoiceHistoryControl _invoiceHistoryControl;
    private readonly SettingsControl _settingsControl;
    private readonly ExpenseControl _expenseControl;
    private readonly RingkasanControl _summaryControl;

    public MainForm(DatabaseService database)
    {
        _database = database;
        _newHaulControl = new NewHaulControl(database);
        _haulListControl = new HaulListControl(database);
        var excelExporter = new ExcelExportService();
        _invoiceControl = new ExcelOutputControl(database, excelExporter, ExcelOutputKind.Invoice);
        _ledgerControl = new ExcelOutputControl(database, excelExporter, ExcelOutputKind.TruckLedger);
        _invoiceHistoryControl = new InvoiceHistoryControl(database, excelExporter);
        _settingsControl = new SettingsControl(database);
        _expenseControl = new ExpenseControl(database);
        _summaryControl = new RingkasanControl(database);
        _invoiceControl.InvoiceGenerated += (_, _) =>
        {
            _invoiceHistoryControl.ReloadData();
            _summaryControl.ReloadData();
        };
        _invoiceHistoryControl.DataChanged += (_, _) =>
        {
            _invoiceControl.ReloadData();
            _summaryControl.ReloadData();
        };
        _settingsControl.SettingsSaved += (_, _) => _invoiceControl.RefreshSettings();
        _settingsControl.LegacyDataImported += (_, result) =>
        {
            _haulListControl.ReloadData();
            _invoiceControl.RefreshSuggestions();
            _invoiceControl.ReloadData();
            _invoiceHistoryControl.ReloadData();
            _expenseControl.RefreshSuggestions();
            _expenseControl.ReloadData();
            _ledgerControl.RefreshSuggestions();
            _ledgerControl.ReloadData();
            _newHaulControl.RefreshSuggestions();
            if (result.LatestDate.HasValue)
            {
                _summaryControl.ShowMonth(result.LatestDate.Value);
            }
            else
            {
                _summaryControl.ReloadData();
            }
            UpdateDatabaseStatus();
        };
        _expenseControl.ExpensesChanged += (_, _) =>
        {
            _ledgerControl.ReloadData();
            _summaryControl.ReloadData();
        };
        _newHaulControl.HaulStored += (_, _) =>
        {
            _haulListControl.ReloadData();
            _expenseControl.RefreshSuggestions();
            _summaryControl.ReloadData();
            UpdateDatabaseStatus();
        };
        _haulListControl.EditRequested += record =>
        {
            ShowPage("Input Angkutan");
            _newHaulControl.LoadRecord(record);
        };
        _haulListControl.DataChanged += (_, _) =>
        {
            _newHaulControl.RefreshSuggestions();
            _summaryControl.ReloadData();
            UpdateDatabaseStatus();
        };

        Text = "Truck Excel Assistant";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(1080, 700);
        Size = new Size(1320, 820);
        BackColor = AppTheme.WindowBackground;
        ForeColor = AppTheme.TextPrimary;
        AutoScaleMode = AutoScaleMode.Dpi;

        BuildLayout();
        UpdateDatabaseStatus();
        ShowPage("Ringkasan");
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.WindowBackground,
            ColumnCount = 2,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220F));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        Controls.Add(root);

        root.Controls.Add(BuildSidebar(), 0, 0);
        root.Controls.Add(BuildMainArea(), 1, 0);
    }

    private Control BuildSidebar()
    {
        var sidebar = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.Sidebar,
            Padding = new Padding(12, 14, 12, 14),
            Margin = Padding.Empty,
            ColumnCount = 1,
            RowCount = 3
        };
        sidebar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        sidebar.RowStyles.Add(new RowStyle(SizeType.Absolute, 88F));
        sidebar.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        sidebar.RowStyles.Add(new RowStyle(SizeType.Absolute, 52F));

        var brand = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.Sidebar,
            Margin = Padding.Empty
        };
        brand.Controls.Add(new Label
        {
            AutoSize = true,
            Text = "TRUCK EXCEL\r\nASSISTANT",
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            Location = new Point(10, 5)
        });
        brand.Controls.Add(new Label
        {
            AutoSize = true,
            Text = "Pembukuan & invoice",
            ForeColor = Color.FromArgb(174, 202, 196),
            Font = new Font("Segoe UI", 8F),
            Location = new Point(10, 51)
        });
        _databaseStatus.Dock = DockStyle.Fill;
        _databaseStatus.ForeColor = Color.FromArgb(196, 211, 208);
        _databaseStatus.BackColor = Color.FromArgb(31, 73, 65);
        _databaseStatus.Font = new Font("Segoe UI", 8F);
        _databaseStatus.Padding = new Padding(10, 7, 6, 6);
        _databaseStatus.TextAlign = ContentAlignment.MiddleLeft;
        var navigation = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoScroll = false,
            Padding = new Padding(0, 12, 0, 0),
            Margin = Padding.Empty
        };

        AddNavigationButton(navigation, "Ringkasan");
        AddNavigationButton(navigation, "Input Angkutan");
        AddNavigationButton(navigation, "Data Angkutan");
        AddNavigationButton(navigation, "Buat Invoice");
        AddNavigationButton(navigation, "Riwayat Invoice");
        AddNavigationButton(navigation, "Pembukuan Truk");
        AddNavigationButton(navigation, "Pengeluaran");
        AddNavigationButton(navigation, "Pengaturan");
        sidebar.Controls.Add(brand, 0, 0);
        sidebar.Controls.Add(navigation, 0, 1);
        sidebar.Controls.Add(_databaseStatus, 0, 2);

        return sidebar;
    }

    private void AddNavigationButton(FlowLayoutPanel navigation, string pageName)
    {
        var button = new Button
        {
            Text = pageName,
            Width = 196,
            Height = 42,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(12, 0, 0, 0),
            Margin = new Padding(0, 0, 0, 4),
            FlatStyle = FlatStyle.Flat,
            BackColor = AppTheme.Sidebar,
            ForeColor = Color.FromArgb(219, 231, 228),
            Cursor = Cursors.Hand,
            TabStop = true,
            UseVisualStyleBackColor = false
        };
        button.FlatAppearance.BorderSize = 0;
        button.FlatAppearance.MouseOverBackColor = AppTheme.SidebarHover;
        button.Click += (_, _) => ShowPage(pageName);
        _navigationButtons[pageName] = button;
        navigation.Controls.Add(button);
    }

    private Control BuildMainArea()
    {
        var main = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Margin = Padding.Empty,
            Padding = Padding.Empty,
            BackColor = AppTheme.WindowBackground
        };
        main.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F));
        main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var header = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.Surface,
            Padding = new Padding(24, 0, 24, 0),
            Margin = Padding.Empty
        };

        _pageTitle.AutoSize = true;
        _pageTitle.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
        _pageTitle.ForeColor = AppTheme.TextPrimary;
        _pageTitle.Location = new Point(24, 19);
        header.Controls.Add(_pageTitle);

        var state = new Label
        {
            AutoSize = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Text = "Data tersimpan secara lokal",
            ForeColor = AppTheme.TextSecondary,
            Font = new Font("Segoe UI", 8.5F),
            Location = new Point(header.Width - 190, 21)
        };
        header.Controls.Add(state);
        header.Resize += (_, _) => state.Left = Math.Max(24, header.ClientSize.Width - state.Width - 24);

        _contentHost.Dock = DockStyle.Fill;
        _contentHost.BackColor = AppTheme.WindowBackground;
        _contentHost.Padding = Padding.Empty;
        _contentHost.Margin = Padding.Empty;

        main.Controls.Add(header, 0, 0);
        main.Controls.Add(_contentHost, 0, 1);
        return main;
    }

    private void ShowPage(string pageName)
    {
        _pageTitle.Text = pageName;
        foreach (var pair in _navigationButtons)
        {
            var selected = pair.Key.Equals(pageName, StringComparison.OrdinalIgnoreCase);
            pair.Value.BackColor = selected ? AppTheme.SidebarHover : AppTheme.Sidebar;
            pair.Value.ForeColor = selected ? Color.White : Color.FromArgb(219, 231, 228);
            pair.Value.Font = new Font("Segoe UI", 9F, selected ? FontStyle.Bold : FontStyle.Regular);
        }

        _contentHost.Controls.Clear();
        Control page;
        switch (pageName)
        {
            case "Ringkasan":
                _summaryControl.ReloadData();
                page = _summaryControl;
                break;
            case "Input Angkutan":
                _newHaulControl.RefreshSuggestions();
                page = _newHaulControl;
                break;
            case "Data Angkutan":
                _haulListControl.ReloadData();
                page = _haulListControl;
                break;
            case "Buat Invoice":
                _invoiceControl.RefreshSuggestions();
                _invoiceControl.RefreshSettings();
                _invoiceControl.ReloadData();
                page = _invoiceControl;
                break;
            case "Riwayat Invoice":
                _invoiceHistoryControl.ReloadData();
                page = _invoiceHistoryControl;
                break;
            case "Pembukuan Truk":
                _ledgerControl.RefreshSuggestions();
                _ledgerControl.ReloadData();
                page = _ledgerControl;
                break;
            case "Pengeluaran":
                _expenseControl.RefreshSuggestions();
                _expenseControl.ReloadData();
                page = _expenseControl;
                break;
            case "Pengaturan":
                _settingsControl.LoadSettings();
                page = _settingsControl;
                break;
            default:
                page = BuildPlaceholder(pageName);
                break;
        }
        page.Dock = DockStyle.Fill;
        _contentHost.Controls.Add(page);
    }

    private void UpdateDatabaseStatus()
    {
        var count = _database.CountHauls();
        _databaseStatus.Text = $"DATABASE LOKAL\r\n{count} perjalanan tersimpan";
        _databaseStatus.AccessibleDescription = _database.DatabasePath;
    }

    private static Control BuildPlaceholder(string pageName)
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.WindowBackground,
            Padding = new Padding(30)
        };
        var title = new Label
        {
            AutoSize = true,
            Text = pageName,
            Font = new Font("Segoe UI", 18F, FontStyle.Bold),
            ForeColor = AppTheme.TextPrimary,
            Location = new Point(30, 30)
        };
        var note = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(650, 0),
            Text = "Halaman ini sudah disiapkan dalam navigasi dan akan dibangun pada tahap berikutnya.",
            Font = new Font("Segoe UI", 10F),
            ForeColor = AppTheme.TextSecondary,
            Location = new Point(33, 75)
        };
        panel.Controls.Add(title);
        panel.Controls.Add(note);
        return panel;
    }
}
