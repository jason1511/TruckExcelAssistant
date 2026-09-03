using System.Diagnostics;
using TruckExcelAssistant.Models;
using TruckExcelAssistant.Services;

namespace TruckExcelAssistant;

public enum ExcelOutputKind
{
    Invoice,
    TruckLedger
}

public sealed class ExcelOutputControl : UserControl
{
    private readonly DatabaseService _database;
    private readonly ExcelExportService _exporter;
    private readonly ExcelOutputKind _kind;
    private readonly DateTimePicker _from = new();
    private readonly DateTimePicker _to = new();
    private readonly DateTimePicker _issueDate = new();
    private readonly ComboBox _subject = new();
    private readonly ComboBox _layout = new();
    private readonly TextBox _invoiceNumber = new();
    private readonly Label _resultCount = new();
    private readonly Label _selectionInfo = new();
    private readonly DataGridView _grid = new();
    private readonly Button _exportButton = AppTheme.CreatePrimaryButton("Buat file Excel");
    private IReadOnlyList<HaulRecord> _records = [];

    public ExcelOutputControl(DatabaseService database, ExcelExportService exporter, ExcelOutputKind kind)
    {
        _database = database;
        _exporter = exporter;
        _kind = kind;
        Dock = DockStyle.Fill;
        BackColor = AppTheme.WindowBackground;
        ForeColor = AppTheme.TextPrimary;
        Padding = new Padding(22, 18, 22, 22);
        AutoScaleMode = AutoScaleMode.Dpi;

        ConfigureInputs();
        BuildLayout();
        ConfigureGrid();
        WireEvents();
    }

    public event EventHandler? InvoiceGenerated;

    public void ReloadData()
    {
        _records = _database.GetSavedHaulsForExport(
            _from.Value.Date,
            _to.Value.Date,
            _kind == ExcelOutputKind.Invoice);
        var subject = _subject.Text.Trim();
        if (!string.IsNullOrWhiteSpace(subject))
        {
            _records = _kind == ExcelOutputKind.Invoice
                ? _records.Where(item => item.Draft.Customer.Equals(subject, StringComparison.OrdinalIgnoreCase)).ToList()
                : _records.Where(item => item.Draft.LicencePlate.Equals(subject, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        _grid.Rows.Clear();
        foreach (var record in _records)
        {
            var draft = record.Draft;
            var row = _grid.Rows.Add(
                true,
                draft.Date.ToString("dd/MM/yyyy"),
                draft.LicencePlate,
                draft.Customer,
                $"{draft.Origin} → {draft.Destination}",
                draft.Cargo,
                $"{IndonesianNumber.Format(draft.ReceivedWeightKg)} kg",
                IndonesianNumber.Rupiah(draft.GrossAmount));
            _grid.Rows[row].Tag = record;
        }
        _resultCount.Text = $"{_records.Count} data ditemukan";
        UpdateSelectionInfo();
    }

    public void RefreshSuggestions()
    {
        var field = _kind == ExcelOutputKind.Invoice ? SuggestionField.Customer : SuggestionField.LicencePlate;
        var current = _subject.Text;
        _subject.BeginUpdate();
        _subject.Items.Clear();
        _subject.Items.AddRange(_database.GetSuggestions(field).Cast<object>().ToArray());
        _subject.EndUpdate();
        _subject.Text = current;
    }

    public void RefreshSettings()
    {
        if (_kind != ExcelOutputKind.Invoice)
        {
            return;
        }
        var settings = _database.GetSettings();
        _layout.SelectedIndex = settings.DefaultInvoiceLayout == OutputLayout.CompactInvoice ? 0 : 1;
        UpdateAutomaticInvoiceNumber();
    }

    private void ConfigureInputs()
    {
        _from.Format = DateTimePickerFormat.Custom;
        _from.CustomFormat = "dd/MM/yyyy";
        _from.Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        _to.Format = DateTimePickerFormat.Custom;
        _to.CustomFormat = "dd/MM/yyyy";
        _to.Value = DateTime.Today;
        _issueDate.Format = DateTimePickerFormat.Custom;
        _issueDate.CustomFormat = "dd/MM/yyyy";
        _issueDate.Value = DateTime.Today;

        _subject.DropDownStyle = ComboBoxStyle.DropDown;
        _subject.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
        _subject.AutoCompleteSource = AutoCompleteSource.ListItems;
        _subject.IntegralHeight = false;
        _subject.MaxDropDownItems = 12;

        _layout.DropDownStyle = ComboBoxStyle.DropDownList;
        _layout.Items.AddRange(["Invoice ringkas (maks. 19 baris)", "Invoice lengkap (maks. 13 baris)"]);
        _layout.SelectedIndex = 1;
        _invoiceNumber.ReadOnly = true;
        _invoiceNumber.BackColor = Color.FromArgb(242, 245, 248);
        _invoiceNumber.TabStop = false;
        UpdateAutomaticInvoiceNumber();

        RefreshSuggestions();
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
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, _kind == ExcelOutputKind.Invoice ? 166F : 94F));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 54F));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        root.Controls.Add(BuildHeading(), 0, 0);
        root.Controls.Add(BuildFilters(), 0, 1);
        root.Controls.Add(BuildActions(), 0, 2);

        var gridPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.Surface,
            BorderStyle = BorderStyle.FixedSingle,
            Margin = Padding.Empty
        };
        gridPanel.Controls.Add(_grid);
        root.Controls.Add(gridPanel, 0, 3);
        Controls.Add(root);
    }

    private Control BuildHeading()
    {
        var panel = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
        panel.Controls.Add(new Label
        {
            Text = _kind == ExcelOutputKind.Invoice ? "Buat invoice Excel" : "Buat pembukuan truk",
            AutoSize = true,
            Font = new Font("Segoe UI", 17F, FontStyle.Bold),
            ForeColor = AppTheme.TextPrimary,
            Location = new Point(0, 0)
        });
        panel.Controls.Add(new Label
        {
            Text = _kind == ExcelOutputKind.Invoice
                ? "Pilih customer, periode, layout, dan perjalanan yang akan masuk invoice."
                : "Pilih periode dan perjalanan. Setiap nomor polisi dibuat sebagai sheet tersendiri.",
            AutoSize = true,
            Font = new Font("Segoe UI", 9F),
            ForeColor = AppTheme.TextSecondary,
            Location = new Point(2, 38)
        });
        return panel;
    }

    private Control BuildFilters()
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.Surface,
            BorderStyle = BorderStyle.FixedSingle,
            Padding = new Padding(14),
            Margin = new Padding(0, 0, 0, 12)
        };
        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            RowCount = _kind == ExcelOutputKind.Invoice ? 2 : 1,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        for (var index = 0; index < 4; index++)
        {
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
        }
        if (_kind == ExcelOutputKind.Invoice)
        {
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
        }
        else
        {
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        }

        AddField(grid, 0, 0, _kind == ExcelOutputKind.Invoice ? "Customer" : "Nomor polisi (kosong = semua)", _subject);
        AddField(grid, 1, 0, "Dari tanggal", _from);
        AddField(grid, 2, 0, "Sampai tanggal", _to);
        var refresh = AppTheme.CreateSecondaryButton("Tampilkan data");
        refresh.Dock = DockStyle.Fill;
        refresh.Margin = new Padding(7, 18, 0, 0);
        refresh.Click += (_, _) => ReloadData();
        grid.Controls.Add(refresh, 3, 0);

        if (_kind == ExcelOutputKind.Invoice)
        {
            AddField(grid, 0, 1, "Layout invoice", _layout, 2);
            AddField(grid, 2, 1, "Nomor invoice (otomatis)", _invoiceNumber);
            AddField(grid, 3, 1, "Tanggal invoice", _issueDate);
        }
        panel.Controls.Add(grid);
        return panel;
    }

    private Control BuildActions()
    {
        var actions = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 5,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = new Padding(0, 0, 0, 10)
        };
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 155F));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190F));

        var all = AppTheme.CreateSecondaryButton("Pilih semua");
        all.Width = 110;
        all.Margin = new Padding(0, 0, 8, 0);
        all.Click += (_, _) => SetAllChecked(true);
        var none = AppTheme.CreateSecondaryButton("Kosongkan");
        none.Width = 100;
        none.Margin = new Padding(0, 0, 12, 0);
        none.Click += (_, _) => SetAllChecked(false);
        _selectionInfo.Dock = DockStyle.Fill;
        _selectionInfo.TextAlign = ContentAlignment.MiddleLeft;
        _selectionInfo.ForeColor = AppTheme.TextSecondary;
        _resultCount.Dock = DockStyle.Fill;
        _resultCount.TextAlign = ContentAlignment.MiddleRight;
        _resultCount.ForeColor = AppTheme.TextSecondary;
        _resultCount.Margin = new Padding(0, 0, 12, 0);
        _exportButton.Dock = DockStyle.Fill;
        _exportButton.Margin = Padding.Empty;
        _exportButton.Click += (_, _) => ExportSelected();

        actions.Controls.Add(all, 0, 0);
        actions.Controls.Add(none, 1, 0);
        actions.Controls.Add(_selectionInfo, 2, 0);
        actions.Controls.Add(_resultCount, 3, 0);
        actions.Controls.Add(_exportButton, 4, 0);
        return actions;
    }

    private void ConfigureGrid()
    {
        _grid.Dock = DockStyle.Fill;
        _grid.BackgroundColor = AppTheme.Surface;
        _grid.BorderStyle = BorderStyle.None;
        _grid.AllowUserToAddRows = false;
        _grid.AllowUserToDeleteRows = false;
        _grid.AllowUserToResizeRows = false;
        _grid.RowHeadersVisible = false;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        _grid.ColumnHeadersHeight = 38;
        _grid.RowTemplate.Height = 36;
        _grid.EnableHeadersVisualStyles = false;
        _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 244, 247);
        _grid.ColumnHeadersDefaultCellStyle.ForeColor = AppTheme.TextPrimary;
        _grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
        _grid.DefaultCellStyle.SelectionBackColor = AppTheme.AccentSoft;
        _grid.DefaultCellStyle.SelectionForeColor = AppTheme.TextPrimary;
        _grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 251, 252);
        _grid.GridColor = Color.FromArgb(231, 235, 240);

        _grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Selected", HeaderText = "Pilih", FillWeight = 35 });
        _grid.Columns.Add("Date", "Tanggal");
        _grid.Columns.Add("Plate", "Nopol");
        _grid.Columns.Add("Customer", "Customer");
        _grid.Columns.Add("Route", "Rute");
        _grid.Columns.Add("Cargo", "Muatan");
        _grid.Columns.Add("Weight", "Berat diterima");
        _grid.Columns.Add("Amount", "Jumlah");
        for (var index = 1; index < _grid.Columns.Count; index++)
        {
            _grid.Columns[index].ReadOnly = true;
        }
    }

    private void WireEvents()
    {
        _from.ValueChanged += (_, _) => EnsureDateOrder();
        _to.ValueChanged += (_, _) => EnsureDateOrder();
        _subject.SelectionChangeCommitted += (_, _) => ReloadData();
        _layout.SelectedIndexChanged += (_, _) => UpdateSelectionInfo();
        _issueDate.ValueChanged += (_, _) => UpdateAutomaticInvoiceNumber();
        _grid.CurrentCellDirtyStateChanged += (_, _) =>
        {
            if (_grid.IsCurrentCellDirty)
            {
                _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        };
        _grid.CellValueChanged += (_, eventArgs) =>
        {
            if (eventArgs.ColumnIndex == 0)
            {
                UpdateSelectionInfo();
            }
        };
    }

    private void ExportSelected()
    {
        var selected = SelectedRecords();
        if (selected.Count == 0)
        {
            MessageBox.Show("Pilih setidaknya satu perjalanan.", "Belum ada data", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        if (_kind == ExcelOutputKind.Invoice && string.IsNullOrWhiteSpace(_subject.Text))
        {
            MessageBox.Show("Isi atau pilih customer terlebih dahulu.", "Customer diperlukan", MessageBoxButtons.OK, MessageBoxIcon.Information);
            _subject.Focus();
            return;
        }
        var suggested = BuildSuggestedFileName();
        var settings = _database.GetSettings();
        var preferredDirectory = Directory.Exists(settings.DefaultExportDirectory)
            ? settings.DefaultExportDirectory
            : _exporter.ExportDirectory;
        using var dialog = new SaveFileDialog
        {
            Title = "Simpan file Excel",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            DefaultExt = "xlsx",
            AddExtension = true,
            InitialDirectory = preferredDirectory,
            FileName = suggested,
            OverwritePrompt = true
        };
        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        try
        {
            if (_kind == ExcelOutputKind.TruckLedger)
            {
                _exporter.ExportTruckLedger(selected, dialog.FileName);
            }
            else if (_layout.SelectedIndex == 0)
            {
                _exporter.ExportCompactInvoice(selected, _subject.Text, _invoiceNumber.Text, _issueDate.Value, dialog.FileName, settings);
                RecordInvoice(selected, OutputLayout.CompactInvoice, dialog.FileName);
            }
            else
            {
                _exporter.ExportCompleteInvoice(selected, _subject.Text, _invoiceNumber.Text, _issueDate.Value, dialog.FileName, settings);
                RecordInvoice(selected, OutputLayout.CompleteInvoice, dialog.FileName);
            }

            var result = MessageBox.Show(
                $"File berhasil dibuat:\n{dialog.FileName}\n\nBuka file sekarang?",
                "Excel selesai dibuat",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);
            if (result == DialogResult.Yes)
            {
                Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
            }
            if (_kind == ExcelOutputKind.Invoice)
            {
                InvoiceGenerated?.Invoke(this, EventArgs.Empty);
                UpdateAutomaticInvoiceNumber();
                ReloadData();
            }
        }
        catch (Exception exception)
        {
            MessageBox.Show(exception.Message, "Gagal membuat Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private IReadOnlyList<HaulRecord> SelectedRecords() =>
        _grid.Rows.Cast<DataGridViewRow>()
            .Where(row => Convert.ToBoolean(row.Cells[0].Value ?? false))
            .Select(row => (HaulRecord)row.Tag!)
            .OrderBy(record => record.Draft.Date)
            .ThenBy(record => record.Id)
            .ToList();

    private void RecordInvoice(IReadOnlyList<HaulRecord> selected, OutputLayout layout, string filePath)
    {
        var total = layout == OutputLayout.CompactInvoice
            ? selected.Sum(item => item.Draft.GrossAmount - item.Draft.BonSangu)
            : selected.Sum(item => item.Draft.GrossAmount + item.Draft.RejectionCost - item.Draft.ClaimAmount);
        _database.RecordGeneratedInvoice(
            _invoiceNumber.Text,
            _issueDate.Value,
            _subject.Text,
            layout,
            total,
            filePath,
            selected.Select(item => item.Id).ToList());
    }

    private void UpdateAutomaticInvoiceNumber()
    {
        if (_kind == ExcelOutputKind.Invoice)
        {
            _invoiceNumber.Text = _database.GetNextInvoiceNumber(_issueDate.Value.Date);
        }
    }

    private void SetAllChecked(bool value)
    {
        foreach (DataGridViewRow row in _grid.Rows)
        {
            row.Cells[0].Value = value;
        }
        UpdateSelectionInfo();
    }

    private void UpdateSelectionInfo()
    {
        var selected = _grid.Rows.Cast<DataGridViewRow>().Count(row => Convert.ToBoolean(row.Cells[0].Value ?? false));
        if (_kind == ExcelOutputKind.Invoice)
        {
            var maximum = _layout.SelectedIndex == 0 ? 19 : 13;
            _selectionInfo.Text = $"Dipilih {selected} dari maksimal {maximum} baris";
            _selectionInfo.ForeColor = selected > maximum ? AppTheme.Warning : AppTheme.TextSecondary;
        }
        else
        {
            _selectionInfo.Text = $"Dipilih {selected} perjalanan • maks. 100 per nopol";
            _selectionInfo.ForeColor = AppTheme.TextSecondary;
        }
        _exportButton.Enabled = selected > 0;
    }

    private void EnsureDateOrder()
    {
        if (_from.Value.Date > _to.Value.Date)
        {
            _to.Value = _from.Value.Date;
        }
    }

    private string BuildSuggestedFileName()
    {
        var subject = string.IsNullOrWhiteSpace(_subject.Text) ? "semua-truk" : _subject.Text.Trim();
        var safe = new string(subject.Select(character => Path.GetInvalidFileNameChars().Contains(character) ? '-' : character).ToArray());
        return _kind == ExcelOutputKind.Invoice
            ? $"Invoice-{safe}-{_invoiceNumber.Text.Trim()}.xlsx"
            : $"Pembukuan-Truk-{safe}-{_from.Value:yyyyMMdd}-{_to.Value:yyyyMMdd}.xlsx";
    }

    private static void AddField(TableLayoutPanel grid, int column, int row, string label, Control input, int columnSpan = 1)
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Margin = new Padding(column == 0 ? 0 : 7, 0, column + columnSpan == 4 ? 0 : 7, 6),
            Padding = Padding.Empty
        };
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 18F));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        panel.Controls.Add(new Label
        {
            Text = label,
            Dock = DockStyle.Fill,
            ForeColor = AppTheme.TextSecondary,
            Font = new Font("Segoe UI", 8.5F)
        }, 0, 0);
        input.Dock = DockStyle.Fill;
        input.Margin = Padding.Empty;
        panel.Controls.Add(input, 0, 1);
        grid.Controls.Add(panel, column, row);
        if (columnSpan > 1)
        {
            grid.SetColumnSpan(panel, columnSpan);
        }
    }
}
