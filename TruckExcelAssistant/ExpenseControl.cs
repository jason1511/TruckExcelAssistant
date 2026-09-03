using TruckExcelAssistant.Models;
using TruckExcelAssistant.Services;

namespace TruckExcelAssistant;

public sealed class ExpenseControl : UserControl
{
    private readonly DatabaseService _database;
    private readonly DateTimePicker _date = new();
    private readonly ComboBox _plate = new();
    private readonly ComboBox _category = new();
    private readonly TextBox _amount = new();
    private readonly TextBox _description = new();
    private readonly TextBox _search = new();
    private readonly ComboBox _filter = new();
    private readonly Label _resultSummary = new();
    private readonly Label _formState = new();
    private readonly DataGridView _grid = new();
    private readonly Button _saveButton = AppTheme.CreatePrimaryButton("Simpan pengeluaran");
    private readonly Button _clearButton = AppTheme.CreateSecondaryButton("Kosongkan form");
    private readonly Button _editButton = AppTheme.CreateSecondaryButton("Edit pengeluaran");
    private readonly Button _trashButton = AppTheme.CreateSecondaryButton("Pindahkan ke Sampah");
    private long? _editingId;

    public ExpenseControl(DatabaseService database)
    {
        _database = database;
        Dock = DockStyle.Fill;
        BackColor = AppTheme.WindowBackground;
        ForeColor = AppTheme.TextPrimary;
        Padding = new Padding(22, 18, 22, 22);
        AutoScaleMode = AutoScaleMode.Dpi;
        ConfigureInputs();
        BuildLayout();
        ConfigureGrid();
        WireEvents();
        RefreshSuggestions();
        UpdateActionButtons();
    }

    public event EventHandler? ExpensesChanged;

    public void RefreshSuggestions()
    {
        var current = _plate.Text;
        _plate.BeginUpdate();
        _plate.Items.Clear();
        _plate.Items.AddRange(_database.GetSuggestions(SuggestionField.LicencePlate).Cast<object>().ToArray());
        _plate.EndUpdate();
        _plate.Text = current;
    }

    public void ReloadData()
    {
        var deletedOnly = _filter.SelectedIndex == 1;
        var records = _database.GetExpenses(_search.Text, deletedOnly);
        _grid.Rows.Clear();
        foreach (var expense in records)
        {
            var row = _grid.Rows.Add(
                expense.Id,
                expense.Date.ToString("dd/MM/yyyy"),
                expense.LicencePlate,
                expense.Category,
                expense.Description,
                IndonesianNumber.Rupiah(expense.Amount),
                expense.DeletedAt is null ? "Aktif" : "Sampah");
            _grid.Rows[row].Tag = expense;
        }
        var total = records.Sum(item => item.Amount);
        _resultSummary.Text = $"{records.Count} data • {IndonesianNumber.Rupiah(total)}";
        UpdateActionButtons();
    }

    private void ConfigureInputs()
    {
        _date.Format = DateTimePickerFormat.Custom;
        _date.CustomFormat = "dd/MM/yyyy";
        _plate.DropDownStyle = ComboBoxStyle.DropDown;
        _plate.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
        _plate.AutoCompleteSource = AutoCompleteSource.ListItems;
        _category.DropDownStyle = ComboBoxStyle.DropDownList;
        _category.Items.AddRange([
            "Bahan bakar",
            "Ban",
            "Servis & suku cadang",
            "Tol & parkir",
            "Pajak & administrasi",
            "Uang jalan tambahan",
            "Lainnya"
        ]);
        _category.SelectedIndex = 2;
        _amount.Text = "0";
        _amount.TextAlign = HorizontalAlignment.Right;
        _amount.BorderStyle = BorderStyle.FixedSingle;
        _description.BorderStyle = BorderStyle.FixedSingle;
        _description.PlaceholderText = "Contoh: ganti dua ban belakang";
        _plate.MaxDropDownItems = 12;
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 72F));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 158F));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64F));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56F));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        root.Controls.Add(BuildHeading(), 0, 0);
        root.Controls.Add(BuildEntryPanel(), 0, 1);
        root.Controls.Add(BuildToolbar(), 0, 2);
        root.Controls.Add(BuildActions(), 0, 3);
        var gridPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.Surface,
            BorderStyle = BorderStyle.FixedSingle,
            Margin = Padding.Empty
        };
        gridPanel.Controls.Add(_grid);
        root.Controls.Add(gridPanel, 0, 4);
        Controls.Add(root);
    }

    private Control BuildHeading()
    {
        var panel = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
        panel.Controls.Add(new Label
        {
            Text = "Pengeluaran truk",
            AutoSize = true,
            Font = new Font("Segoe UI", 17F, FontStyle.Bold),
            ForeColor = AppTheme.TextPrimary,
            Location = new Point(0, 0)
        });
        panel.Controls.Add(new Label
        {
            Text = "Catat biaya per nomor polisi agar otomatis masuk ke Pembukuan Truk.",
            AutoSize = true,
            Font = new Font("Segoe UI", 9F),
            ForeColor = AppTheme.TextSecondary,
            Location = new Point(2, 38)
        });
        return panel;
    }

    private Control BuildEntryPanel()
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
            ColumnCount = 5,
            RowCount = 2,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 18F));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22F));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22F));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 18F));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
        AddField(grid, 0, 0, "Tanggal", _date);
        AddField(grid, 1, 0, "Nomor polisi", _plate);
        AddField(grid, 2, 0, "Kategori", _category);
        AddField(grid, 3, 0, "Jumlah (Rp)", _amount);
        _saveButton.Dock = DockStyle.Fill;
        _saveButton.Margin = new Padding(7, 19, 0, 6);
        grid.Controls.Add(_saveButton, 4, 0);
        AddField(grid, 0, 1, "Keterangan", _description, 4);
        _clearButton.Dock = DockStyle.Fill;
        _clearButton.Margin = new Padding(7, 19, 0, 6);
        grid.Controls.Add(_clearButton, 4, 1);
        panel.Controls.Add(grid);
        return panel;
    }

    private Control BuildToolbar()
    {
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
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 210F));
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110F));
        _search.Dock = DockStyle.Fill;
        _search.PlaceholderText = "Cari nopol, kategori, atau keterangan...";
        _search.Margin = new Padding(0, 0, 12, 0);
        _filter.Dock = DockStyle.Fill;
        _filter.DropDownStyle = ComboBoxStyle.DropDownList;
        _filter.Items.AddRange(["Pengeluaran aktif", "Sampah"]);
        _filter.SelectedIndex = 0;
        _filter.Margin = new Padding(0, 0, 12, 0);
        _resultSummary.Dock = DockStyle.Fill;
        _resultSummary.TextAlign = ContentAlignment.MiddleRight;
        _resultSummary.ForeColor = AppTheme.TextSecondary;
        _resultSummary.Margin = new Padding(0, 0, 12, 0);
        var reload = AppTheme.CreateSecondaryButton("Muat ulang");
        reload.Dock = DockStyle.Fill;
        reload.Margin = Padding.Empty;
        reload.Click += (_, _) => ReloadData();
        toolbar.Controls.Add(_search, 0, 0);
        toolbar.Controls.Add(_filter, 1, 0);
        toolbar.Controls.Add(_resultSummary, 2, 0);
        toolbar.Controls.Add(reload, 3, 0);
        return toolbar;
    }

    private Control BuildActions()
    {
        var actions = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = new Padding(0, 0, 0, 10)
        };
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 155F));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 185F));
        _formState.Dock = DockStyle.Fill;
        _formState.Text = "Pilih baris untuk mengedit atau memindahkannya ke Sampah.";
        _formState.ForeColor = AppTheme.TextSecondary;
        _formState.TextAlign = ContentAlignment.MiddleLeft;
        _editButton.Dock = DockStyle.Fill;
        _editButton.Margin = new Padding(0, 0, 8, 0);
        _trashButton.Dock = DockStyle.Fill;
        _trashButton.Margin = Padding.Empty;
        actions.Controls.Add(_formState, 0, 0);
        actions.Controls.Add(_editButton, 1, 0);
        actions.Controls.Add(_trashButton, 2, 0);
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
        _grid.Columns.Add("Id", "ID");
        _grid.Columns.Add("Date", "Tanggal");
        _grid.Columns.Add("Plate", "Nopol");
        _grid.Columns.Add("Category", "Kategori");
        _grid.Columns.Add("Description", "Keterangan");
        _grid.Columns.Add("Amount", "Jumlah");
        _grid.Columns.Add("Status", "Status");
        _grid.Columns[0].Visible = false;
    }

    private void WireEvents()
    {
        _saveButton.Click += (_, _) => SaveExpense();
        _clearButton.Click += (_, _) => ResetForm();
        _editButton.Click += (_, _) => EditSelected();
        _trashButton.Click += (_, _) => TrashOrRestoreSelected();
        _search.TextChanged += (_, _) => ReloadData();
        _filter.SelectedIndexChanged += (_, _) => ReloadData();
        _grid.SelectionChanged += (_, _) => UpdateActionButtons();
        _grid.CellDoubleClick += (_, eventArgs) =>
        {
            if (eventArgs.RowIndex >= 0)
            {
                EditSelected();
            }
        };
        _amount.Leave += (_, _) =>
        {
            if (IndonesianNumber.TryParse(_amount.Text, out var value))
            {
                _amount.Text = IndonesianNumber.Format(value);
            }
        };
    }

    private void SaveExpense()
    {
        if (string.IsNullOrWhiteSpace(_plate.Text))
        {
            MessageBox.Show("Isi atau pilih nomor polisi.", "Nomor polisi diperlukan", MessageBoxButtons.OK, MessageBoxIcon.Information);
            _plate.Focus();
            return;
        }
        if (!IndonesianNumber.TryParse(_amount.Text, out var amount) || amount <= 0)
        {
            MessageBox.Show("Jumlah pengeluaran harus lebih besar dari nol.", "Jumlah tidak valid", MessageBoxButtons.OK, MessageBoxIcon.Information);
            _amount.Focus();
            return;
        }
        var category = _category.SelectedItem?.ToString() ?? "Lainnya";
        if (_editingId is long id)
        {
            _database.UpdateExpense(id, _date.Value, _plate.Text, category, _description.Text, amount);
        }
        else
        {
            _database.AddExpense(_date.Value, _plate.Text, category, _description.Text, amount);
        }
        ResetForm();
        RefreshSuggestions();
        ReloadData();
        ExpensesChanged?.Invoke(this, EventArgs.Empty);
    }

    private ExpenseRecord? SelectedExpense() =>
        _grid.SelectedRows.Count == 1 ? _grid.SelectedRows[0].Tag as ExpenseRecord : null;

    private void EditSelected()
    {
        var expense = SelectedExpense();
        if (expense is null || expense.DeletedAt is not null)
        {
            return;
        }
        _editingId = expense.Id;
        _date.Value = expense.Date;
        _plate.Text = expense.LicencePlate;
        _category.SelectedItem = expense.Category;
        if (_category.SelectedIndex < 0)
        {
            _category.SelectedItem = "Lainnya";
        }
        _amount.Text = IndonesianNumber.Format(expense.Amount);
        _description.Text = expense.Description;
        _saveButton.Text = "Simpan perubahan";
        _clearButton.Text = "Batal mengedit";
        _formState.Text = $"MENGEDIT PENGELUARAN #{expense.Id}";
        _formState.ForeColor = AppTheme.Accent;
        _plate.Focus();
    }

    private void TrashOrRestoreSelected()
    {
        var expense = SelectedExpense();
        if (expense is null)
        {
            return;
        }
        if (expense.DeletedAt is not null)
        {
            _database.RestoreExpenseFromTrash(expense.Id);
        }
        else
        {
            var result = MessageBox.Show(
                $"Pindahkan pengeluaran {IndonesianNumber.Rupiah(expense.Amount)} ke Sampah?",
                "Pindahkan ke Sampah",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (result != DialogResult.Yes)
            {
                return;
            }
            _database.MoveExpenseToTrash(expense.Id);
            if (_editingId == expense.Id)
            {
                ResetForm();
            }
        }
        ReloadData();
        ExpensesChanged?.Invoke(this, EventArgs.Empty);
    }

    private void ResetForm()
    {
        _editingId = null;
        _date.Value = DateTime.Today;
        _plate.Text = string.Empty;
        _category.SelectedIndex = 2;
        _amount.Text = "0";
        _description.Clear();
        _saveButton.Text = "Simpan pengeluaran";
        _clearButton.Text = "Kosongkan form";
        _formState.Text = "Pilih baris untuk mengedit atau memindahkannya ke Sampah.";
        _formState.ForeColor = AppTheme.TextSecondary;
    }

    private void UpdateActionButtons()
    {
        var expense = SelectedExpense();
        var isTrash = expense?.DeletedAt is not null;
        _editButton.Enabled = expense is not null && !isTrash;
        _trashButton.Enabled = expense is not null;
        _trashButton.Text = isTrash ? "Pulihkan" : "Pindahkan ke Sampah";
    }

    private static void AddField(TableLayoutPanel grid, int column, int row, string label, Control input, int span = 1)
    {
        var field = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Margin = new Padding(column == 0 ? 0 : 7, 0, column + span == grid.ColumnCount ? 0 : 7, 6),
            Padding = Padding.Empty
        };
        field.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
        field.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        field.Controls.Add(new Label
        {
            Text = label,
            Dock = DockStyle.Fill,
            ForeColor = AppTheme.TextSecondary,
            Font = new Font("Segoe UI", 8.5F)
        }, 0, 0);
        input.Dock = DockStyle.Fill;
        input.Margin = Padding.Empty;
        field.Controls.Add(input, 0, 1);
        grid.Controls.Add(field, column, row);
        if (span > 1)
        {
            grid.SetColumnSpan(field, span);
        }
    }
}
