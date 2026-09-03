using TruckExcelAssistant.Models;
using TruckExcelAssistant.Services;

namespace TruckExcelAssistant;

public sealed class RingkasanControl : UserControl
{
    private readonly DatabaseService _database;
    private readonly DateTimePicker _month = new();
    private readonly Label _periodNote = new();
    private readonly Label _revenueValue = CreateKpiValueLabel();
    private readonly Label _revenueNote = CreateKpiNoteLabel();
    private readonly Label _expenseValue = CreateKpiValueLabel();
    private readonly Label _expenseNote = CreateKpiNoteLabel();
    private readonly Label _netValue = CreateKpiValueLabel();
    private readonly Label _netNote = CreateKpiNoteLabel();
    private readonly Label _invoiceValue = CreateKpiValueLabel();
    private readonly Label _invoiceNote = CreateKpiNoteLabel();
    private readonly Label _truckCount = new();
    private readonly Label _invoiceCount = new();
    private readonly DataGridView _truckGrid = new();
    private readonly DataGridView _invoiceGrid = new();

    public RingkasanControl(DatabaseService database)
    {
        _database = database;
        Dock = DockStyle.Fill;
        BackColor = AppTheme.WindowBackground;
        ForeColor = AppTheme.TextPrimary;
        Padding = new Padding(22, 18, 22, 22);
        AutoScaleMode = AutoScaleMode.Dpi;

        BuildLayout();
        ConfigureTruckGrid();
        ConfigureInvoiceGrid();
        _month.ValueChanged += (_, _) => ReloadData();
    }

    public void ReloadData()
    {
        var summary = _database.GetDashboardSummary(_month.Value);
        _periodNote.Text = $"Aktivitas {summary.Month:MMMM yyyy}";
        _revenueValue.Text = IndonesianNumber.Rupiah(summary.Revenue);
        _revenueNote.Text = summary.HaulCount == 1 ? "1 perjalanan tersimpan" : $"{summary.HaulCount} perjalanan tersimpan";
        _expenseValue.Text = IndonesianNumber.Rupiah(summary.TotalExpenses);
        _expenseNote.Text = $"Angkutan {IndonesianNumber.Rupiah(summary.EmbeddedExpenses)}  •  Lainnya {IndonesianNumber.Rupiah(summary.StandaloneExpenses)}";
        _netValue.Text = IndonesianNumber.Rupiah(summary.Net);
        _netValue.ForeColor = summary.Net < 0 ? AppTheme.Warning : AppTheme.Accent;
        _netNote.Text = "Pemasukan dikurangi seluruh pengeluaran";
        _invoiceValue.Text = IndonesianNumber.Rupiah(summary.OutstandingInvoiceAmount);
        _invoiceNote.Text = summary.OutstandingInvoiceCount == 1
            ? "1 invoice belum lunas (semua periode)"
            : $"{summary.OutstandingInvoiceCount} invoice belum lunas (semua periode)";

        _truckGrid.Rows.Clear();
        foreach (var truck in summary.Trucks.OrderByDescending(row => row.Net))
        {
            var rowIndex = _truckGrid.Rows.Add(
                truck.LicencePlate,
                truck.HaulCount,
                IndonesianNumber.Rupiah(truck.Revenue),
                IndonesianNumber.Rupiah(truck.Expenses),
                IndonesianNumber.Rupiah(truck.Net));
            if (truck.Net < 0)
            {
                _truckGrid.Rows[rowIndex].Cells[4].Style.ForeColor = AppTheme.Warning;
            }
        }
        _truckCount.Text = summary.Trucks.Count == 1 ? "1 truk" : $"{summary.Trucks.Count} truk";

        _invoiceGrid.Rows.Clear();
        foreach (var invoice in summary.RecentInvoices)
        {
            var rowIndex = _invoiceGrid.Rows.Add(
                invoice.InvoiceNumber,
                invoice.InvoiceDate.ToString("dd/MM/yyyy"),
                invoice.Customer,
                IndonesianNumber.Rupiah(invoice.TotalAmount),
                StatusText(invoice.Status));
            if (invoice.Status == InvoiceStatus.Generated)
            {
                _invoiceGrid.Rows[rowIndex].Cells[4].Style.ForeColor = AppTheme.Warning;
            }
        }
        _invoiceCount.Text = summary.RecentInvoices.Count == 1 ? "1 invoice" : $"{summary.RecentInvoices.Count} invoice terbaru";
    }

    public void ShowMonth(DateTime date)
    {
        if (_month.Value.Year == date.Year && _month.Value.Month == date.Month)
        {
            ReloadData();
            return;
        }
        _month.Value = new DateTime(date.Year, date.Month, 1);
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
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 72F));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 144F));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        root.Controls.Add(BuildHeading(), 0, 0);
        root.Controls.Add(BuildKpis(), 0, 1);
        root.Controls.Add(BuildTables(), 0, 2);
        Controls.Add(root);
    }

    private Control BuildHeading()
    {
        var heading = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        heading.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        heading.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 270F));

        var text = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
        text.Controls.Add(new Label
        {
            Text = "Ringkasan usaha",
            AutoSize = true,
            Font = new Font("Segoe UI", 17F, FontStyle.Bold),
            ForeColor = AppTheme.TextPrimary,
            Location = new Point(0, 0)
        });
        _periodNote.AutoSize = true;
        _periodNote.Font = new Font("Segoe UI", 9F);
        _periodNote.ForeColor = AppTheme.TextSecondary;
        _periodNote.Location = new Point(2, 38);
        text.Controls.Add(_periodNote);

        var period = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = new Padding(0, 3, 0, 17)
        };
        period.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        period.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92F));
        _month.Dock = DockStyle.Fill;
        _month.Format = DateTimePickerFormat.Custom;
        _month.CustomFormat = "MMMM yyyy";
        _month.ShowUpDown = true;
        _month.Margin = new Padding(0, 0, 8, 0);
        var refresh = AppTheme.CreateSecondaryButton("Muat ulang");
        refresh.Dock = DockStyle.Fill;
        refresh.Margin = Padding.Empty;
        refresh.Click += (_, _) => ReloadData();
        period.Controls.Add(_month, 0, 0);
        period.Controls.Add(refresh, 1, 0);
        heading.Controls.Add(text, 0, 0);
        heading.Controls.Add(period, 1, 0);
        return heading;
    }

    private Control BuildKpis()
    {
        var cards = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = new Padding(0, 0, 0, 14)
        };
        for (var index = 0; index < 4; index++)
        {
            cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
        }
        cards.Controls.Add(CreateKpiCard("PEMASUKAN ANGKUTAN", _revenueValue, _revenueNote, new Padding(0, 0, 8, 0)), 0, 0);
        cards.Controls.Add(CreateKpiCard("TOTAL PENGELUARAN", _expenseValue, _expenseNote, new Padding(4, 0, 4, 0)), 1, 0);
        cards.Controls.Add(CreateKpiCard("HASIL BERSIH", _netValue, _netNote, new Padding(4, 0, 4, 0)), 2, 0);
        cards.Controls.Add(CreateKpiCard("INVOICE BELUM LUNAS", _invoiceValue, _invoiceNote, new Padding(8, 0, 0, 0)), 3, 0);
        return cards;
    }

    private Control BuildTables()
    {
        var tables = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        tables.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58F));
        tables.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42F));
        tables.Controls.Add(CreateGridCard("Kinerja per truk", "Pemasukan, pengeluaran, dan hasil bersih", _truckCount, _truckGrid, new Padding(0, 0, 7, 0)), 0, 0);
        tables.Controls.Add(CreateGridCard("Invoice terbaru", "Enam invoice terakhir dari semua periode", _invoiceCount, _invoiceGrid, new Padding(7, 0, 0, 0)), 1, 0);
        return tables;
    }

    private static Control CreateKpiCard(string title, Label value, Label note, Padding margin)
    {
        var card = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.Surface,
            BorderStyle = BorderStyle.FixedSingle,
            Margin = margin,
            Padding = new Padding(16, 13, 16, 10)
        };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Margin = Padding.Empty };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 24F));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        layout.Controls.Add(new Label
        {
            Text = title,
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 8F, FontStyle.Bold),
            ForeColor = AppTheme.TextSecondary,
            TextAlign = ContentAlignment.MiddleLeft
        }, 0, 0);
        layout.Controls.Add(value, 0, 1);
        layout.Controls.Add(note, 0, 2);
        card.Controls.Add(layout);
        return card;
    }

    private static Control CreateGridCard(string title, string subtitle, Label count, DataGridView grid, Padding margin)
    {
        var card = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.Surface,
            BorderStyle = BorderStyle.FixedSingle,
            ColumnCount = 1,
            RowCount = 2,
            Margin = margin,
            Padding = Padding.Empty
        };
        card.RowStyles.Add(new RowStyle(SizeType.Absolute, 64F));
        card.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        var header = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty, Padding = Padding.Empty };
        header.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            ForeColor = AppTheme.TextPrimary,
            Location = new Point(14, 10)
        });
        header.Controls.Add(new Label
        {
            Text = subtitle,
            AutoSize = true,
            Font = new Font("Segoe UI", 8F),
            ForeColor = AppTheme.TextSecondary,
            Location = new Point(15, 34)
        });
        count.AutoSize = true;
        count.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        count.Font = new Font("Segoe UI", 8F);
        count.ForeColor = AppTheme.TextSecondary;
        count.Location = new Point(header.Width - 70, 14);
        header.Controls.Add(count);
        header.Resize += (_, _) => count.Left = Math.Max(15, header.ClientSize.Width - count.Width - 14);
        card.Controls.Add(header, 0, 0);
        card.Controls.Add(grid, 0, 1);
        return card;
    }

    private void ConfigureTruckGrid()
    {
        ConfigureGrid(_truckGrid);
        _truckGrid.Columns.Add("Plate", "Nopol");
        _truckGrid.Columns.Add("Hauls", "Perjalanan");
        _truckGrid.Columns.Add("Revenue", "Pemasukan");
        _truckGrid.Columns.Add("Expenses", "Pengeluaran");
        _truckGrid.Columns.Add("Net", "Hasil bersih");
        _truckGrid.Columns[0].FillWeight = 78;
        _truckGrid.Columns[1].FillWeight = 62;
        _truckGrid.Columns[2].FillWeight = 100;
        _truckGrid.Columns[3].FillWeight = 100;
        _truckGrid.Columns[4].FillWeight = 100;
    }

    private void ConfigureInvoiceGrid()
    {
        ConfigureGrid(_invoiceGrid);
        _invoiceGrid.Columns.Add("Number", "Nomor");
        _invoiceGrid.Columns.Add("Date", "Tanggal");
        _invoiceGrid.Columns.Add("Customer", "Customer");
        _invoiceGrid.Columns.Add("Total", "Total");
        _invoiceGrid.Columns.Add("Status", "Status");
        _invoiceGrid.Columns[0].FillWeight = 92;
        _invoiceGrid.Columns[1].FillWeight = 66;
        _invoiceGrid.Columns[2].FillWeight = 105;
        _invoiceGrid.Columns[3].FillWeight = 90;
        _invoiceGrid.Columns[4].FillWeight = 62;
    }

    private static void ConfigureGrid(DataGridView grid)
    {
        grid.Dock = DockStyle.Fill;
        grid.BackgroundColor = AppTheme.Surface;
        grid.BorderStyle = BorderStyle.None;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.AllowUserToResizeRows = false;
        grid.ReadOnly = true;
        grid.MultiSelect = false;
        grid.RowHeadersVisible = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        grid.ColumnHeadersHeight = 38;
        grid.RowTemplate.Height = 38;
        grid.EnableHeadersVisualStyles = false;
        grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 244, 247);
        grid.ColumnHeadersDefaultCellStyle.ForeColor = AppTheme.TextPrimary;
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
        grid.DefaultCellStyle.SelectionBackColor = AppTheme.AccentSoft;
        grid.DefaultCellStyle.SelectionForeColor = AppTheme.TextPrimary;
        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 251, 252);
        grid.GridColor = Color.FromArgb(231, 235, 240);
    }

    private static Label CreateKpiValueLabel() => new()
    {
        Dock = DockStyle.Fill,
        Font = new Font("Segoe UI", 15F, FontStyle.Bold),
        ForeColor = AppTheme.TextPrimary,
        TextAlign = ContentAlignment.MiddleLeft,
        AutoEllipsis = true
    };

    private static Label CreateKpiNoteLabel() => new()
    {
        Dock = DockStyle.Fill,
        Font = new Font("Segoe UI", 7.8F),
        ForeColor = AppTheme.TextSecondary,
        TextAlign = ContentAlignment.TopLeft,
        AutoEllipsis = true
    };

    private static string StatusText(InvoiceStatus status) => status switch
    {
        InvoiceStatus.Generated => "Belum lunas",
        InvoiceStatus.Paid => "Lunas",
        InvoiceStatus.Cancelled => "Dibatalkan",
        _ => status.ToString()
    };
}
