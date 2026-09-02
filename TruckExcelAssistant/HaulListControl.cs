using TruckExcelAssistant.Models;
using TruckExcelAssistant.Services;

namespace TruckExcelAssistant;

public sealed class HaulListControl : UserControl
{
    private readonly DatabaseService _database;
    private readonly TextBox _search = new();
    private readonly Label _resultCount = new();
    private readonly DataGridView _grid = new();

    public HaulListControl(DatabaseService database)
    {
        _database = database;
        Dock = DockStyle.Fill;
        BackColor = AppTheme.WindowBackground;
        ForeColor = AppTheme.TextPrimary;
        Padding = new Padding(22, 18, 22, 22);
        AutoScaleMode = AutoScaleMode.Dpi;

        BuildLayout();
        ConfigureGrid();
        WireEvents();
    }

    public void ReloadData()
    {
        var records = _database.GetHauls(_search.Text);
        _grid.Rows.Clear();
        foreach (var record in records)
        {
            var draft = record.Draft;
            _grid.Rows.Add(
                record.Id,
                draft.Date.ToString("dd/MM/yyyy"),
                draft.LicencePlate,
                draft.Customer,
                BuildRoute(draft),
                draft.Cargo,
                $"{IndonesianNumber.Format(draft.ReceivedWeightKg)} kg",
                IndonesianNumber.Rupiah(draft.RatePerKg),
                IndonesianNumber.Rupiah(draft.GrossAmount),
                record.Status == HaulStatus.Draft ? "Draft" : "Tersimpan");
        }

        _resultCount.Text = records.Count == 1
            ? "1 perjalanan"
            : $"{records.Count} perjalanan";
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 70F));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 62F));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var heading = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
        heading.Controls.Add(new Label
        {
            Text = "Data angkutan",
            AutoSize = true,
            Font = new Font("Segoe UI", 17F, FontStyle.Bold),
            ForeColor = AppTheme.TextPrimary,
            Location = new Point(0, 0)
        });
        heading.Controls.Add(new Label
        {
            Text = "Cari perjalanan yang sudah tersimpan atau masih berupa draft.",
            AutoSize = true,
            Font = new Font("Segoe UI", 9F),
            ForeColor = AppTheme.TextSecondary,
            Location = new Point(2, 38)
        });

        var toolbar = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.Surface,
            BorderStyle = BorderStyle.FixedSingle,
            ColumnCount = 3,
            RowCount = 1,
            Margin = new Padding(0, 0, 0, 12),
            Padding = new Padding(12, 9, 12, 9)
        };
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110F));

        _search.Dock = DockStyle.Fill;
        _search.BorderStyle = BorderStyle.FixedSingle;
        _search.PlaceholderText = "Cari nopol, customer, muatan, atau lokasi...";
        _search.Margin = new Padding(0, 0, 12, 0);

        _resultCount.Dock = DockStyle.Fill;
        _resultCount.TextAlign = ContentAlignment.MiddleRight;
        _resultCount.ForeColor = AppTheme.TextSecondary;
        _resultCount.Margin = new Padding(0, 0, 12, 0);

        var refreshButton = AppTheme.CreateSecondaryButton("Muat ulang");
        refreshButton.Dock = DockStyle.Fill;
        refreshButton.Height = 30;
        refreshButton.Margin = Padding.Empty;
        refreshButton.Click += (_, _) => ReloadData();

        toolbar.Controls.Add(_search, 0, 0);
        toolbar.Controls.Add(_resultCount, 1, 0);
        toolbar.Controls.Add(refreshButton, 2, 0);

        var gridPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.Surface,
            BorderStyle = BorderStyle.FixedSingle,
            Margin = Padding.Empty
        };
        gridPanel.Controls.Add(_grid);

        root.Controls.Add(heading, 0, 0);
        root.Controls.Add(toolbar, 0, 1);
        root.Controls.Add(gridPanel, 0, 2);
        Controls.Add(root);
    }

    private void ConfigureGrid()
    {
        _grid.Dock = DockStyle.Fill;
        _grid.BackgroundColor = AppTheme.Surface;
        _grid.BorderStyle = BorderStyle.None;
        _grid.AllowUserToAddRows = false;
        _grid.AllowUserToDeleteRows = false;
        _grid.AllowUserToResizeRows = false;
        _grid.ReadOnly = true;
        _grid.MultiSelect = false;
        _grid.RowHeadersVisible = false;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        _grid.ColumnHeadersHeight = 38;
        _grid.RowTemplate.Height = 38;
        _grid.EnableHeadersVisualStyles = false;
        _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 244, 247);
        _grid.ColumnHeadersDefaultCellStyle.ForeColor = AppTheme.TextPrimary;
        _grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
        _grid.DefaultCellStyle.BackColor = AppTheme.Surface;
        _grid.DefaultCellStyle.ForeColor = AppTheme.TextPrimary;
        _grid.DefaultCellStyle.SelectionBackColor = AppTheme.AccentSoft;
        _grid.DefaultCellStyle.SelectionForeColor = AppTheme.TextPrimary;
        _grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 251, 252);
        _grid.GridColor = Color.FromArgb(231, 235, 240);

        _grid.Columns.Add("Id", "ID");
        _grid.Columns.Add("Date", "Tanggal");
        _grid.Columns.Add("Plate", "Nopol");
        _grid.Columns.Add("Customer", "Customer");
        _grid.Columns.Add("Route", "Rute");
        _grid.Columns.Add("Cargo", "Muatan");
        _grid.Columns.Add("Weight", "Berat diterima");
        _grid.Columns.Add("Rate", "Ongkos");
        _grid.Columns.Add("Gross", "Jumlah");
        _grid.Columns.Add("Status", "Status");
        _grid.Columns[0].Visible = false;
        _grid.Columns[1].FillWeight = 72;
        _grid.Columns[2].FillWeight = 82;
        _grid.Columns[3].FillWeight = 105;
        _grid.Columns[4].FillWeight = 120;
        _grid.Columns[5].FillWeight = 90;
        _grid.Columns[6].FillWeight = 85;
        _grid.Columns[7].FillWeight = 80;
        _grid.Columns[8].FillWeight = 90;
        _grid.Columns[9].FillWeight = 72;
    }

    private void WireEvents()
    {
        _search.TextChanged += (_, _) => ReloadData();
    }

    private static string BuildRoute(HaulDraft draft)
    {
        if (string.IsNullOrWhiteSpace(draft.Origin) && string.IsNullOrWhiteSpace(draft.Destination))
        {
            return "—";
        }
        return $"{draft.Origin} → {draft.Destination}";
    }
}
