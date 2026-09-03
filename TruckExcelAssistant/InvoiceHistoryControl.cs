using System.Diagnostics;
using TruckExcelAssistant.Models;
using TruckExcelAssistant.Services;

namespace TruckExcelAssistant;

public sealed class InvoiceHistoryControl : UserControl
{
    private readonly DatabaseService _database;
    private readonly ExcelExportService _exporter;
    private readonly TextBox _search = new();
    private readonly ComboBox _statusFilter = new();
    private readonly Label _resultCount = new();
    private readonly DataGridView _grid = new();
    private readonly Button _openButton = AppTheme.CreateSecondaryButton("Buka Excel");
    private readonly Button _folderButton = AppTheme.CreateSecondaryButton("Buka folder");
    private readonly Button _regenerateButton = AppTheme.CreateSecondaryButton("Buat ulang");
    private readonly Button _paidButton = AppTheme.CreateSecondaryButton("Tandai lunas");
    private readonly Button _cancelButton = AppTheme.CreateSecondaryButton("Batalkan invoice");

    public InvoiceHistoryControl(DatabaseService database, ExcelExportService exporter)
    {
        _database = database;
        _exporter = exporter;
        Dock = DockStyle.Fill;
        BackColor = AppTheme.WindowBackground;
        ForeColor = AppTheme.TextPrimary;
        Padding = new Padding(22, 18, 22, 22);
        AutoScaleMode = AutoScaleMode.Dpi;

        BuildLayout();
        ConfigureGrid();
        WireEvents();
        UpdateActionButtons();
    }

    public event EventHandler? DataChanged;

    public void ReloadData()
    {
        var status = _statusFilter.SelectedIndex switch
        {
            1 => InvoiceStatus.Generated,
            2 => InvoiceStatus.Paid,
            3 => InvoiceStatus.Cancelled,
            _ => (InvoiceStatus?)null
        };
        var invoices = _database.GetInvoices(_search.Text, status);
        _grid.Rows.Clear();
        foreach (var invoice in invoices)
        {
            var exists = File.Exists(invoice.FilePath);
            var row = _grid.Rows.Add(
                invoice.InvoiceNumber,
                invoice.InvoiceDate.ToString("dd/MM/yyyy"),
                invoice.Customer,
                LayoutText(invoice.Layout),
                invoice.HaulCount,
                IndonesianNumber.Rupiah(invoice.TotalAmount),
                StatusText(invoice.Status),
                exists ? "Tersedia" : "Tidak ditemukan");
            _grid.Rows[row].Tag = invoice;
            if (!exists)
            {
                _grid.Rows[row].Cells[7].Style.ForeColor = AppTheme.Warning;
            }
        }
        _resultCount.Text = invoices.Count == 1 ? "1 invoice" : $"{invoices.Count} invoice";
        UpdateActionButtons();
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 72F));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64F));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58F));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var heading = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
        heading.Controls.Add(new Label
        {
            Text = "Riwayat invoice",
            AutoSize = true,
            Font = new Font("Segoe UI", 17F, FontStyle.Bold),
            ForeColor = AppTheme.TextPrimary,
            Location = new Point(0, 0)
        });
        heading.Controls.Add(new Label
        {
            Text = "Buka kembali, buat ulang, tandai lunas, atau batalkan invoice yang pernah dibuat.",
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
            ColumnCount = 4,
            RowCount = 1,
            Margin = new Padding(0, 0, 0, 12),
            Padding = new Padding(12, 9, 12, 9)
        };
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170F));
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 125F));
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110F));
        _search.Dock = DockStyle.Fill;
        _search.PlaceholderText = "Cari nomor invoice, customer, atau nama file...";
        _search.Margin = new Padding(0, 0, 12, 0);
        _statusFilter.Dock = DockStyle.Fill;
        _statusFilter.DropDownStyle = ComboBoxStyle.DropDownList;
        _statusFilter.Items.AddRange(["Semua status", "Belum lunas", "Lunas", "Dibatalkan"]);
        _statusFilter.SelectedIndex = 0;
        _statusFilter.Margin = new Padding(0, 0, 12, 0);
        _resultCount.Dock = DockStyle.Fill;
        _resultCount.TextAlign = ContentAlignment.MiddleRight;
        _resultCount.ForeColor = AppTheme.TextSecondary;
        _resultCount.Margin = new Padding(0, 0, 12, 0);
        var refresh = AppTheme.CreateSecondaryButton("Muat ulang");
        refresh.Dock = DockStyle.Fill;
        refresh.Margin = Padding.Empty;
        refresh.Click += (_, _) => ReloadData();
        toolbar.Controls.Add(_search, 0, 0);
        toolbar.Controls.Add(_statusFilter, 1, 0);
        toolbar.Controls.Add(_resultCount, 2, 0);
        toolbar.Controls.Add(refresh, 3, 0);

        var actions = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 6,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = new Padding(0, 0, 0, 10)
        };
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115F));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115F));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130F));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 145F));
        actions.Controls.Add(new Label
        {
            Text = "Pilih satu invoice atau klik dua kali untuk membuka Excel.",
            Dock = DockStyle.Fill,
            ForeColor = AppTheme.TextSecondary,
            TextAlign = ContentAlignment.MiddleLeft
        }, 0, 0);
        AddAction(actions, _openButton, 1);
        AddAction(actions, _folderButton, 2);
        AddAction(actions, _regenerateButton, 3);
        AddAction(actions, _paidButton, 4);
        AddAction(actions, _cancelButton, 5, true);

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
        root.Controls.Add(actions, 0, 2);
        root.Controls.Add(gridPanel, 0, 3);
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
        _grid.DefaultCellStyle.SelectionBackColor = AppTheme.AccentSoft;
        _grid.DefaultCellStyle.SelectionForeColor = AppTheme.TextPrimary;
        _grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 251, 252);
        _grid.GridColor = Color.FromArgb(231, 235, 240);
        _grid.Columns.Add("Number", "Nomor invoice");
        _grid.Columns.Add("Date", "Tanggal");
        _grid.Columns.Add("Customer", "Customer");
        _grid.Columns.Add("Layout", "Layout");
        _grid.Columns.Add("Hauls", "Perjalanan");
        _grid.Columns.Add("Total", "Total");
        _grid.Columns.Add("Status", "Status");
        _grid.Columns.Add("File", "File Excel");
    }

    private void WireEvents()
    {
        _search.TextChanged += (_, _) => ReloadData();
        _statusFilter.SelectedIndexChanged += (_, _) => ReloadData();
        _grid.SelectionChanged += (_, _) => UpdateActionButtons();
        _grid.CellDoubleClick += (_, eventArgs) =>
        {
            if (eventArgs.RowIndex >= 0)
            {
                OpenSelected();
            }
        };
        _openButton.Click += (_, _) => OpenSelected();
        _folderButton.Click += (_, _) => OpenFolder();
        _regenerateButton.Click += (_, _) => RegenerateSelected();
        _paidButton.Click += (_, _) => TogglePaidStatus();
        _cancelButton.Click += (_, _) => CancelSelected();
    }

    private InvoiceRecord? SelectedInvoice() =>
        _grid.SelectedRows.Count == 1 ? _grid.SelectedRows[0].Tag as InvoiceRecord : null;

    private void OpenSelected()
    {
        var invoice = SelectedInvoice();
        if (invoice is null)
        {
            return;
        }
        if (!File.Exists(invoice.FilePath))
        {
            var result = MessageBox.Show(
                "File Excel tidak ditemukan di lokasi terakhir. Buat ulang invoice sekarang?",
                "File tidak ditemukan",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                Regenerate(invoice);
            }
            return;
        }
        Process.Start(new ProcessStartInfo(invoice.FilePath) { UseShellExecute = true });
    }

    private void OpenFolder()
    {
        var invoice = SelectedInvoice();
        if (invoice is null)
        {
            return;
        }
        var folder = Path.GetDirectoryName(invoice.FilePath);
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            MessageBox.Show("Folder penyimpanan tidak ditemukan.", "Folder tidak tersedia", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        Process.Start(new ProcessStartInfo(folder) { UseShellExecute = true });
    }

    private void RegenerateSelected()
    {
        var invoice = SelectedInvoice();
        if (invoice is not null)
        {
            Regenerate(invoice);
        }
    }

    private void Regenerate(InvoiceRecord invoice)
    {
        var hauls = _database.GetInvoiceHauls(invoice.Id);
        if (hauls.Count == 0)
        {
            MessageBox.Show("Invoice ini tidak memiliki data perjalanan yang dapat dibuat ulang.", "Data tidak tersedia", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        var previousFolder = Path.GetDirectoryName(invoice.FilePath);
        using var dialog = new SaveFileDialog
        {
            Title = "Buat ulang invoice Excel",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            DefaultExt = "xlsx",
            AddExtension = true,
            OverwritePrompt = true,
            InitialDirectory = !string.IsNullOrWhiteSpace(previousFolder) && Directory.Exists(previousFolder)
                ? previousFolder
                : _exporter.ExportDirectory,
            FileName = string.IsNullOrWhiteSpace(Path.GetFileName(invoice.FilePath))
                ? $"Invoice-{invoice.InvoiceNumber}.xlsx"
                : Path.GetFileName(invoice.FilePath)
        };
        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }
        try
        {
            if (invoice.Layout == OutputLayout.CompactInvoice)
            {
                _exporter.ExportCompactInvoice(hauls, invoice.Customer, invoice.InvoiceNumber, invoice.InvoiceDate, dialog.FileName);
            }
            else
            {
                _exporter.ExportCompleteInvoice(hauls, invoice.Customer, invoice.InvoiceNumber, invoice.InvoiceDate, dialog.FileName);
            }
            _database.UpdateInvoiceFilePath(invoice.Id, dialog.FileName);
            ReloadData();
            MessageBox.Show("Invoice berhasil dibuat ulang.", "Excel selesai dibuat", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception exception)
        {
            MessageBox.Show(exception.Message, "Gagal membuat ulang invoice", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void TogglePaidStatus()
    {
        var invoice = SelectedInvoice();
        if (invoice is null || invoice.Status == InvoiceStatus.Cancelled)
        {
            return;
        }
        var next = invoice.Status == InvoiceStatus.Paid ? InvoiceStatus.Generated : InvoiceStatus.Paid;
        _database.UpdateInvoiceStatus(invoice.Id, next);
        ReloadData();
        DataChanged?.Invoke(this, EventArgs.Empty);
    }

    private void CancelSelected()
    {
        var invoice = SelectedInvoice();
        if (invoice is null || invoice.Status == InvoiceStatus.Cancelled)
        {
            return;
        }
        var result = MessageBox.Show(
            $"Batalkan {invoice.InvoiceNumber}?\n\nPerjalanan di dalamnya akan dapat dipilih untuk invoice baru. File Excel lama tidak akan dihapus.",
            "Batalkan invoice",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);
        if (result != DialogResult.Yes)
        {
            return;
        }
        _database.UpdateInvoiceStatus(invoice.Id, InvoiceStatus.Cancelled);
        ReloadData();
        DataChanged?.Invoke(this, EventArgs.Empty);
    }

    private void UpdateActionButtons()
    {
        var invoice = SelectedInvoice();
        var selected = invoice is not null;
        _openButton.Enabled = selected;
        _folderButton.Enabled = selected;
        _regenerateButton.Enabled = selected;
        _paidButton.Enabled = selected && invoice!.Status != InvoiceStatus.Cancelled;
        _cancelButton.Enabled = selected && invoice!.Status != InvoiceStatus.Cancelled;
        _paidButton.Text = invoice?.Status == InvoiceStatus.Paid ? "Tandai belum lunas" : "Tandai lunas";
    }

    private static void AddAction(TableLayoutPanel panel, Button button, int column, bool last = false)
    {
        button.Dock = DockStyle.Fill;
        button.Margin = new Padding(0, 0, last ? 0 : 8, 0);
        panel.Controls.Add(button, column, 0);
    }

    private static string LayoutText(OutputLayout layout) => layout == OutputLayout.CompactInvoice
        ? "Ringkas"
        : "Lengkap";

    private static string StatusText(InvoiceStatus status) => status switch
    {
        InvoiceStatus.Paid => "Lunas",
        InvoiceStatus.Cancelled => "Dibatalkan",
        _ => "Belum lunas"
    };
}
