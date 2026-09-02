using TruckExcelAssistant.Models;
using TruckExcelAssistant.Services;

namespace TruckExcelAssistant;

public sealed class HaulListControl : UserControl
{
    private readonly DatabaseService _database;
    private readonly TextBox _search = new();
    private readonly ComboBox _filter = new();
    private readonly Label _resultCount = new();
    private readonly DataGridView _grid = new();
    private readonly Button _editButton;
    private readonly Button _trashButton;

    public HaulListControl(DatabaseService database)
    {
        _database = database;
        _editButton = AppTheme.CreateSecondaryButton("Edit / lanjutkan");
        _trashButton = AppTheme.CreateSecondaryButton("Pindahkan ke Sampah");

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

    public event Action<HaulRecord>? EditRequested;

    public event EventHandler? DataChanged;

    public void ReloadData()
    {
        var status = _filter.SelectedIndex switch
        {
            1 => HaulStatus.Saved,
            2 => HaulStatus.Draft,
            _ => (HaulStatus?)null
        };
        var deletedOnly = _filter.SelectedIndex == 3;
        var records = _database.GetHauls(_search.Text, status, deletedOnly);

        _grid.Rows.Clear();
        foreach (var record in records)
        {
            var draft = record.Draft;
            var rowIndex = _grid.Rows.Add(
                record.Id,
                draft.Date.ToString("dd/MM/yyyy"),
                draft.LicencePlate,
                draft.Customer,
                BuildRoute(draft),
                draft.Cargo,
                $"{IndonesianNumber.Format(draft.ReceivedWeightKg)} kg",
                IndonesianNumber.Rupiah(draft.RatePerKg),
                IndonesianNumber.Rupiah(draft.GrossAmount),
                record.DeletedAt is not null
                    ? "Sampah"
                    : record.Status == HaulStatus.Draft ? "Draft" : "Tersimpan");
            _grid.Rows[rowIndex].Tag = record;
        }

        _resultCount.Text = records.Count == 1
            ? "1 perjalanan"
            : $"{records.Count} perjalanan";
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
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 70F));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 62F));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 52F));
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
            Text = "Cari, lanjutkan draft, edit, atau pulihkan data dari Sampah.",
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
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 155F));
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 135F));
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105F));

        _search.Dock = DockStyle.Fill;
        _search.BorderStyle = BorderStyle.FixedSingle;
        _search.PlaceholderText = "Cari nopol, customer, muatan, atau lokasi...";
        _search.Margin = new Padding(0, 0, 12, 0);

        _filter.Dock = DockStyle.Fill;
        _filter.DropDownStyle = ComboBoxStyle.DropDownList;
        _filter.Items.AddRange(["Semua data", "Tersimpan", "Draft", "Sampah"]);
        _filter.SelectedIndex = 0;
        _filter.Margin = new Padding(0, 0, 12, 0);

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
        toolbar.Controls.Add(_filter, 1, 0);
        toolbar.Controls.Add(_resultCount, 2, 0);
        toolbar.Controls.Add(refreshButton, 3, 0);

        var actions = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = new Padding(0, 0, 0, 10)
        };
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 185F));
        actions.Controls.Add(new Label
        {
            Text = "Pilih satu baris atau klik dua kali untuk mengedit.",
            Dock = DockStyle.Fill,
            ForeColor = AppTheme.TextSecondary,
            TextAlign = ContentAlignment.MiddleLeft
        }, 0, 0);

        _editButton.Dock = DockStyle.Fill;
        _editButton.Margin = new Padding(0, 0, 8, 0);
        _trashButton.Dock = DockStyle.Fill;
        _trashButton.Margin = Padding.Empty;
        actions.Controls.Add(_editButton, 1, 0);
        actions.Controls.Add(_trashButton, 2, 0);

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
    }

    private void WireEvents()
    {
        _search.TextChanged += (_, _) => ReloadData();
        _filter.SelectedIndexChanged += (_, _) => ReloadData();
        _grid.SelectionChanged += (_, _) => UpdateActionButtons();
        _grid.CellDoubleClick += (_, e) =>
        {
            if (e.RowIndex >= 0)
            {
                EditSelected();
            }
        };
        _editButton.Click += (_, _) => EditSelected();
        _trashButton.Click += (_, _) => TrashOrRestoreSelected();
    }

    private HaulRecord? SelectedRecord() =>
        _grid.SelectedRows.Count == 1
            ? _grid.SelectedRows[0].Tag as HaulRecord
            : null;

    private void EditSelected()
    {
        var record = SelectedRecord();
        if (record is null || record.DeletedAt is not null)
        {
            return;
        }
        EditRequested?.Invoke(record);
    }

    private void TrashOrRestoreSelected()
    {
        var record = SelectedRecord();
        if (record is null)
        {
            return;
        }

        if (record.DeletedAt is not null)
        {
            _database.RestoreFromTrash(record.Id);
        }
        else
        {
            var identity = string.IsNullOrWhiteSpace(record.Draft.LicencePlate)
                ? $"data #{record.Id}"
                : record.Draft.LicencePlate;
            var result = MessageBox.Show(
                $"Pindahkan {identity} ke Sampah? Data dapat dipulihkan kembali.",
                "Pindahkan ke Sampah",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (result != DialogResult.Yes)
            {
                return;
            }
            _database.MoveToTrash(record.Id);
        }

        ReloadData();
        DataChanged?.Invoke(this, EventArgs.Empty);
    }

    private void UpdateActionButtons()
    {
        var record = SelectedRecord();
        var isTrash = record?.DeletedAt is not null;
        _editButton.Enabled = record is not null && !isTrash;
        _trashButton.Enabled = record is not null;
        _trashButton.Text = isTrash ? "Pulihkan" : "Pindahkan ke Sampah";
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
