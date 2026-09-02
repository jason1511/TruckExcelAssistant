using TruckExcelAssistant.Models;
using TruckExcelAssistant.Services;

namespace TruckExcelAssistant;

public sealed class NewHaulControl : UserControl
{
    private readonly DateTimePicker _date = new();
    private readonly ComboBox _licencePlate = new();
    private readonly ComboBox _cargo = new();
    private readonly ComboBox _customer = new();
    private readonly ComboBox _origin = new();
    private readonly ComboBox _destination = new();
    private readonly TextBox _loadedWeight = new();
    private readonly TextBox _receivedWeight = new();
    private readonly TextBox _rate = new();
    private readonly TextBox _grossAmount = new();
    private readonly TextBox _bonSangu = new();
    private readonly TextBox _rejectionCost = new();
    private readonly TextBox _claimAmount = new();
    private readonly TextBox _driverRoadMoney = new();
    private readonly TextBox _otherExpense = new();
    private readonly TextBox _notes = new();

    private readonly Panel _customerField = new();
    private readonly Panel _bonSanguField = new();
    private readonly Panel _rejectionCostField = new();
    private readonly Panel _claimField = new();
    private readonly Panel _driverRoadMoneyField = new();
    private readonly Panel _otherExpenseField = new();
    private readonly Panel _notesField = new();
    private readonly TableLayoutPanel _adjustmentGrid = new();
    private readonly Label _differenceValue = new();
    private readonly Label _grossValue = new();
    private readonly Label _adjustmentLabel = new();
    private readonly Label _adjustmentValue = new();
    private readonly Label _finalLabel = new();
    private readonly Label _finalValue = new();
    private readonly Label _calculationNote = new();
    private readonly DataGridView _previewGrid = new();
    private readonly Dictionary<OutputLayout, Button> _layoutButtons = [];
    private readonly ErrorProvider _errors = new();

    private OutputLayout _layout = OutputLayout.CompleteInvoice;

    public NewHaulControl()
    {
        Dock = DockStyle.Fill;
        BackColor = AppTheme.WindowBackground;
        ForeColor = AppTheme.TextPrimary;
        AutoScaleMode = AutoScaleMode.Dpi;
        Padding = new Padding(22, 18, 22, 22);

        ConfigureInputs();
        BuildLayout();
        WireEvents();
        SetLayout(OutputLayout.CompleteInvoice);
        Recalculate();
    }

    private void ConfigureInputs()
    {
        ConfigureCombo(_licencePlate, "Contoh: N-9103-UZ");
        ConfigureCombo(_cargo, "Contoh: SBM atau Jagung");
        ConfigureCombo(_customer, "Nama customer");
        ConfigureCombo(_origin, "Lokasi muat");
        ConfigureCombo(_destination, "Lokasi bongkar");

        _cargo.Items.AddRange(["SBM", "Jagung", "Tepung", "Pupuk", "Pasir"]);
        _customer.Items.AddRange(["Agrico", "Miguno"]);
        _origin.Items.AddRange(["Lumajang", "Jember", "Surabaya", "Teluk Lamong", "Jakarta", "Gresik"]);
        _destination.Items.AddRange(["Semarang", "Cirebon", "Balaraja", "Surabaya", "Lumajang"]);

        _date.Format = DateTimePickerFormat.Custom;
        _date.CustomFormat = "dd/MM/yyyy";
        _date.Dock = DockStyle.Fill;
        _date.CalendarForeColor = AppTheme.TextPrimary;

        foreach (var textBox in NumericInputs())
        {
            ConfigureTextBox(textBox);
            textBox.TextAlign = HorizontalAlignment.Right;
            textBox.Text = "0";
        }

        _grossAmount.ReadOnly = true;
        _grossAmount.BackColor = Color.FromArgb(242, 245, 248);
        _grossAmount.TabStop = false;

        ConfigureTextBox(_notes);
        _notes.PlaceholderText = "Catatan perjalanan atau biaya";
        _notes.Multiline = false;

        _errors.BlinkStyle = ErrorBlinkStyle.NeverBlink;
        _errors.ContainerControl = this;
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            BackColor = AppTheme.WindowBackground,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 78F));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 480F));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        root.Resize += (_, _) =>
        {
            var availableHeight = Math.Max(0, root.ClientSize.Height - 78);
            root.RowStyles[1].Height = Math.Clamp(availableHeight * 0.56F, 360F, 520F);
        };
        Controls.Add(root);

        root.Controls.Add(BuildHeading(), 0, 0);
        root.Controls.Add(BuildWorkspace(), 0, 1);
        root.Controls.Add(BuildPreview(), 0, 2);
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
        heading.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var titles = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
        titles.Controls.Add(new Label
        {
            Text = "Tambah perjalanan",
            AutoSize = true,
            Font = new Font("Segoe UI", 17F, FontStyle.Bold),
            ForeColor = AppTheme.TextPrimary,
            Location = new Point(0, 0)
        });
        titles.Controls.Add(new Label
        {
            Text = "Masukkan satu kali, gunakan untuk pembukuan dan invoice Excel.",
            AutoSize = true,
            Font = new Font("Segoe UI", 9F),
            ForeColor = AppTheme.TextSecondary,
            Location = new Point(2, 38)
        });

        var layouts = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Margin = new Padding(0, 2, 0, 0),
            Padding = Padding.Empty
        };
        AddLayoutButton(layouts, OutputLayout.TruckLedger, "Pembukuan", 118);
        AddLayoutButton(layouts, OutputLayout.CompactInvoice, "Invoice ringkas", 138);
        AddLayoutButton(layouts, OutputLayout.CompleteInvoice, "Invoice lengkap", 142);

        heading.Controls.Add(titles, 0, 0);
        heading.Controls.Add(layouts, 1, 0);
        return heading;
    }

    private void AddLayoutButton(FlowLayoutPanel host, OutputLayout layout, string text, int width)
    {
        var button = AppTheme.CreateSecondaryButton(text);
        button.Width = width;
        button.Margin = new Padding(4, 0, 0, 0);
        button.Tag = layout;
        button.Click += (_, _) => SetLayout(layout);
        _layoutButtons[layout] = button;
        host.Controls.Add(button);
    }

    private Control BuildWorkspace()
    {
        var workspace = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        workspace.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 76F));
        workspace.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 24F));

        var formScroll = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = AppTheme.Surface,
            BorderStyle = BorderStyle.FixedSingle,
            Padding = new Padding(16),
            Margin = new Padding(0, 0, 14, 14)
        };

        var sections = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        formScroll.Resize += (_, _) =>
        {
            sections.Width = Math.Max(100, formScroll.ClientSize.Width - 35);
            foreach (Control section in sections.Controls)
            {
                section.Width = Math.Max(100, sections.ClientSize.Width - 4);
            }
        };

        sections.Controls.Add(BuildIdentitySection());
        sections.Controls.Add(BuildWeightSection());
        sections.Controls.Add(BuildAdjustmentsSection());
        formScroll.Controls.Add(sections);

        workspace.Controls.Add(formScroll, 0, 0);
        workspace.Controls.Add(BuildSummary(), 1, 0);
        return workspace;
    }

    private Control BuildIdentitySection()
    {
        var fields = CreateFieldGrid(2);
        AddField(fields, 0, 0, "Tanggal", _date);
        AddField(fields, 1, 0, "Nomor polisi", _licencePlate);
        AddField(fields, 2, 0, "Jenis muatan", _cargo);
        _customerField.Controls.Add(CreateField("Customer", _customer));
        ConfigureFieldContainer(_customerField);
        fields.Controls.Add(_customerField, 3, 0);
        AddField(fields, 0, 1, "Dari", _origin, 2);
        AddField(fields, 2, 1, "Tujuan", _destination, 2);
        return CreateSection("Identitas perjalanan", fields);
    }

    private Control BuildWeightSection()
    {
        var fields = CreateFieldGrid(1);
        AddField(fields, 0, 0, "Berat muat (kg)", _loadedWeight);
        AddField(fields, 1, 0, "Berat diterima (kg)", _receivedWeight);
        AddField(fields, 2, 0, "Ongkos (Rp/kg)", _rate);
        AddField(fields, 3, 0, "Jumlah", _grossAmount);
        return CreateSection("Berat dan ongkos", fields);
    }

    private Control BuildAdjustmentsSection()
    {
        ConfigureFieldGrid(_adjustmentGrid, 2);

        ConfigureFieldContainer(_bonSanguField);
        _bonSanguField.Controls.Add(CreateField("Bon sangu", _bonSangu));

        ConfigureFieldContainer(_rejectionCostField);
        _rejectionCostField.Controls.Add(CreateField("Biaya tolakan / uang makan", _rejectionCost));

        ConfigureFieldContainer(_claimField);
        _claimField.Controls.Add(CreateField("Nilai klaim", _claimAmount));

        PrepareDynamicField(_driverRoadMoneyField, "Uang jalan sopir", _driverRoadMoney);
        PrepareDynamicField(_otherExpenseField, "Biaya lainnya", _otherExpense);
        PrepareDynamicField(_notesField, "Keterangan", _notes);

        return CreateSection("Biaya dan penyesuaian", _adjustmentGrid);
    }

    private Control BuildSummary()
    {
        var summary = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.Surface,
            BorderStyle = BorderStyle.FixedSingle,
            Padding = new Padding(16),
            Margin = new Padding(0, 0, 0, 14)
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 2,
            RowCount = 7,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 54F));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 46F));

        var title = new Label
        {
            Text = "Ringkasan",
            AutoSize = true,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            ForeColor = AppTheme.TextPrimary,
            Margin = new Padding(0, 0, 0, 14)
        };
        layout.Controls.Add(title, 0, 0);
        layout.SetColumnSpan(title, 2);
        AddSummaryRow(layout, 1, "Selisih berat", _differenceValue);
        AddSummaryRow(layout, 2, "Jumlah angkutan", _grossValue);
        AddSummaryRow(layout, 3, "Penyesuaian", _adjustmentValue, _adjustmentLabel);
        AddSummaryRow(layout, 4, "Perkiraan bersih", _finalValue, _finalLabel, true);

        _calculationNote.AutoSize = true;
        _calculationNote.MaximumSize = new Size(245, 0);
        _calculationNote.ForeColor = AppTheme.TextSecondary;
        _calculationNote.Font = new Font("Segoe UI", 8F);
        _calculationNote.Margin = new Padding(0, 12, 0, 6);
        layout.Controls.Add(_calculationNote, 0, 5);
        layout.SetColumnSpan(_calculationNote, 2);

        var draftButton = AppTheme.CreateSecondaryButton("Simpan sebagai draft");
        draftButton.Dock = DockStyle.Top;
        draftButton.Click += (_, _) => ShowDraftMessage();
        var saveButton = AppTheme.CreatePrimaryButton("Simpan angkutan");
        saveButton.Dock = DockStyle.Top;
        saveButton.Click += (_, _) => SaveHaul();

        var actions = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            RowCount = 2,
            Margin = new Padding(0, 8, 0, 0),
            Padding = Padding.Empty
        };
        actions.Controls.Add(draftButton, 0, 0);
        actions.Controls.Add(saveButton, 0, 1);
        layout.Controls.Add(actions, 0, 6);
        layout.SetColumnSpan(actions, 2);

        summary.Controls.Add(layout);
        return summary;
    }

    private Control BuildPreview()
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.Surface,
            BorderStyle = BorderStyle.FixedSingle,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };

        var header = new Panel { Dock = DockStyle.Top, Height = 38, Padding = new Padding(12, 0, 12, 0) };
        header.Controls.Add(new Label
        {
            AutoSize = true,
            Text = "Pratinjau baris Excel",
            Font = new Font("Segoe UI", 9F, FontStyle.Bold),
            ForeColor = AppTheme.TextPrimary,
            Location = new Point(12, 11)
        });

        ConfigurePreviewGrid();
        panel.Controls.Add(_previewGrid);
        panel.Controls.Add(header);
        return panel;
    }

    private void ConfigurePreviewGrid()
    {
        _previewGrid.Dock = DockStyle.Fill;
        _previewGrid.BackgroundColor = AppTheme.Surface;
        _previewGrid.BorderStyle = BorderStyle.None;
        _previewGrid.AllowUserToAddRows = false;
        _previewGrid.AllowUserToDeleteRows = false;
        _previewGrid.AllowUserToResizeRows = false;
        _previewGrid.ReadOnly = true;
        _previewGrid.RowHeadersVisible = false;
        _previewGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _previewGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        _previewGrid.ColumnHeadersHeight = 34;
        _previewGrid.RowTemplate.Height = 34;
        _previewGrid.EnableHeadersVisualStyles = false;
        _previewGrid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 244, 247);
        _previewGrid.ColumnHeadersDefaultCellStyle.ForeColor = AppTheme.TextPrimary;
        _previewGrid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
        _previewGrid.DefaultCellStyle.BackColor = AppTheme.Surface;
        _previewGrid.DefaultCellStyle.ForeColor = AppTheme.TextPrimary;
        _previewGrid.DefaultCellStyle.SelectionBackColor = AppTheme.AccentSoft;
        _previewGrid.DefaultCellStyle.SelectionForeColor = AppTheme.TextPrimary;
        _previewGrid.GridColor = Color.FromArgb(231, 235, 240);

        _previewGrid.Columns.Add("Date", "Tanggal");
        _previewGrid.Columns.Add("Plate", "Nopol");
        _previewGrid.Columns.Add("Route", "Rute");
        _previewGrid.Columns.Add("Cargo", "Barang");
        _previewGrid.Columns.Add("Weight", "Berat diterima");
        _previewGrid.Columns.Add("Rate", "Ongkos");
        _previewGrid.Columns.Add("Gross", "Jumlah");
        _previewGrid.Columns.Add("Final", "Hasil akhir");
    }

    private void WireEvents()
    {
        foreach (var textBox in NumericInputs())
        {
            textBox.TextChanged += (_, _) => Recalculate();
            textBox.Leave += (_, _) => FormatNumericTextBox(textBox);
            textBox.KeyPress += NumericTextBoxKeyPress;
        }

        _date.ValueChanged += (_, _) => UpdatePreview();
        _licencePlate.TextChanged += (_, _) =>
        {
            var selectionStart = _licencePlate.SelectionStart;
            _licencePlate.Text = _licencePlate.Text.ToUpperInvariant();
            _licencePlate.SelectionStart = Math.Min(selectionStart, _licencePlate.Text.Length);
            UpdatePreview();
        };
        _cargo.TextChanged += (_, _) => UpdatePreview();
        _customer.TextChanged += (_, _) => UpdatePreview();
        _origin.TextChanged += (_, _) => UpdatePreview();
        _destination.TextChanged += (_, _) => UpdatePreview();
    }

    private void SetLayout(OutputLayout layout)
    {
        _layout = layout;
        foreach (var pair in _layoutButtons)
        {
            var selected = pair.Key == layout;
            pair.Value.BackColor = selected ? AppTheme.Accent : AppTheme.Surface;
            pair.Value.ForeColor = selected ? Color.White : AppTheme.TextPrimary;
            pair.Value.FlatAppearance.BorderColor = selected ? AppTheme.Accent : AppTheme.InputBorder;
        }

        ArrangeAdjustmentFields(layout);

        _adjustmentLabel.Text = layout switch
        {
            OutputLayout.CompactInvoice => "Bon sangu",
            OutputLayout.CompleteInvoice => "Biaya / klaim",
            _ => "Uang jalan / biaya"
        };
        _finalLabel.Text = layout == OutputLayout.TruckLedger ? "Perkiraan bersih" : "Total invoice";
        _calculationNote.Text = layout switch
        {
            OutputLayout.CompactInvoice => "Layout invoice ringkas: jumlah angkutan dikurangi bon sangu. Customer dipilih secara terpisah.",
            OutputLayout.CompleteInvoice => "Layout invoice lengkap: jumlah ditambah biaya tolakan lalu dikurangi klaim. Customer dipilih secara terpisah.",
            _ => "Pembukuan: pemasukan angkutan dikurangi uang jalan sopir dan biaya lainnya."
        };

        Recalculate();
    }

    private void ArrangeAdjustmentFields(OutputLayout layout)
    {
        _adjustmentGrid.SuspendLayout();
        _adjustmentGrid.Controls.Clear();

        switch (layout)
        {
            case OutputLayout.CompactInvoice:
                AddExistingField(_adjustmentGrid, _bonSanguField, 0, 0, 2);
                AddExistingField(_adjustmentGrid, _driverRoadMoneyField, 2, 0);
                AddExistingField(_adjustmentGrid, _otherExpenseField, 3, 0);
                break;
            case OutputLayout.CompleteInvoice:
                AddExistingField(_adjustmentGrid, _rejectionCostField, 0, 0);
                AddExistingField(_adjustmentGrid, _claimField, 1, 0);
                AddExistingField(_adjustmentGrid, _driverRoadMoneyField, 2, 0);
                AddExistingField(_adjustmentGrid, _otherExpenseField, 3, 0);
                break;
            default:
                AddExistingField(_adjustmentGrid, _driverRoadMoneyField, 0, 0, 2);
                AddExistingField(_adjustmentGrid, _otherExpenseField, 2, 0, 2);
                break;
        }

        AddExistingField(_adjustmentGrid, _notesField, 0, 1, 4);
        _adjustmentGrid.ResumeLayout(true);
    }

    private void Recalculate()
    {
        var draft = ReadDraft();
        _grossAmount.Text = IndonesianNumber.Format(draft.GrossAmount);

        _differenceValue.Text = $"{IndonesianNumber.Format(draft.WeightDifferenceKg)} kg";
        _differenceValue.ForeColor = draft.WeightDifferenceKg > 0 ? AppTheme.Warning : AppTheme.TextPrimary;
        _grossValue.Text = IndonesianNumber.Rupiah(draft.GrossAmount);

        var adjustment = _layout switch
        {
            OutputLayout.CompactInvoice => -draft.BonSangu,
            OutputLayout.CompleteInvoice => draft.RejectionCost - draft.ClaimAmount,
            _ => -(draft.DriverRoadMoney + draft.OtherExpense)
        };
        _adjustmentValue.Text = IndonesianNumber.Rupiah(adjustment);
        _finalValue.Text = IndonesianNumber.Rupiah(draft.FinalAmount);
        UpdatePreview(draft);
    }

    private HaulDraft ReadDraft()
    {
        return new HaulDraft(
            _date.Value.Date,
            _licencePlate.Text.Trim().ToUpperInvariant(),
            _cargo.Text.Trim(),
            _customer.Text.Trim(),
            _origin.Text.Trim(),
            _destination.Text.Trim(),
            ReadNumber(_loadedWeight),
            ReadNumber(_receivedWeight),
            ReadNumber(_rate),
            ReadNumber(_bonSangu),
            ReadNumber(_rejectionCost),
            ReadNumber(_claimAmount),
            ReadNumber(_driverRoadMoney),
            ReadNumber(_otherExpense),
            _notes.Text.Trim(),
            _layout);
    }

    private void UpdatePreview() => UpdatePreview(ReadDraft());

    private void UpdatePreview(HaulDraft draft)
    {
        if (_previewGrid.Columns.Count == 0)
        {
            return;
        }

        _previewGrid.Rows.Clear();
        _previewGrid.Rows.Add(
            draft.Date.ToString("dd/MM/yyyy"),
            string.IsNullOrWhiteSpace(draft.LicencePlate) ? "—" : draft.LicencePlate,
            BuildRoute(draft),
            string.IsNullOrWhiteSpace(draft.Cargo) ? "—" : draft.Cargo,
            $"{IndonesianNumber.Format(draft.ReceivedWeightKg)} kg",
            IndonesianNumber.Rupiah(draft.RatePerKg),
            IndonesianNumber.Rupiah(draft.GrossAmount),
            IndonesianNumber.Rupiah(draft.FinalAmount));
    }

    private void SaveHaul()
    {
        _errors.Clear();
        var draft = ReadDraft();
        var isValid = true;

        isValid &= RequireText(_licencePlate, "Nomor polisi wajib diisi.");
        isValid &= RequireText(_cargo, "Jenis muatan wajib diisi.");
        isValid &= RequireText(_customer, "Customer wajib diisi.");
        if (draft.ReceivedWeightKg <= 0)
        {
            _errors.SetError(_receivedWeight, "Berat diterima harus lebih dari nol.");
            isValid = false;
        }
        if (draft.RatePerKg <= 0)
        {
            _errors.SetError(_rate, "Ongkos harus lebih dari nol.");
            isValid = false;
        }

        if (!isValid)
        {
            MessageBox.Show(
                "Periksa kembali kolom yang ditandai.",
                "Data belum lengkap",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        MessageBox.Show(
            $"Angkutan {draft.LicencePlate} siap disimpan.\n\n" +
            "Penyimpanan database akan diaktifkan pada tahap berikutnya.",
            "Validasi berhasil",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private void ShowDraftMessage()
    {
        MessageBox.Show(
            "Form sudah siap untuk fitur draft. Penyimpanan database akan ditambahkan pada tahap berikutnya.",
            "Draft",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private bool RequireText(ComboBox control, string message)
    {
        if (!string.IsNullOrWhiteSpace(control.Text))
        {
            return true;
        }

        _errors.SetError(control, message);
        return false;
    }

    private static string BuildRoute(HaulDraft draft)
    {
        if (string.IsNullOrWhiteSpace(draft.Origin) && string.IsNullOrWhiteSpace(draft.Destination))
        {
            return "—";
        }
        return $"{draft.Origin} → {draft.Destination}";
    }

    private static decimal ReadNumber(TextBox textBox) =>
        IndonesianNumber.TryParse(textBox.Text, out var value) ? value : 0;

    private static void FormatNumericTextBox(TextBox textBox)
    {
        if (IndonesianNumber.TryParse(textBox.Text, out var value))
        {
            textBox.Text = IndonesianNumber.Format(value);
        }
    }

    private static void NumericTextBoxKeyPress(object? sender, KeyPressEventArgs e)
    {
        if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar is not '.' and not ',')
        {
            e.Handled = true;
        }
    }

    private IEnumerable<TextBox> NumericInputs()
    {
        yield return _loadedWeight;
        yield return _receivedWeight;
        yield return _rate;
        yield return _bonSangu;
        yield return _rejectionCost;
        yield return _claimAmount;
        yield return _driverRoadMoney;
        yield return _otherExpense;
    }

    private static void ConfigureCombo(ComboBox comboBox, string placeholder)
    {
        comboBox.Dock = DockStyle.Fill;
        comboBox.DropDownStyle = ComboBoxStyle.DropDown;
        comboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
        comboBox.AutoCompleteSource = AutoCompleteSource.ListItems;
        comboBox.FlatStyle = FlatStyle.Standard;
        comboBox.BackColor = AppTheme.Surface;
        comboBox.ForeColor = AppTheme.TextPrimary;
        comboBox.IntegralHeight = false;
        comboBox.DropDownHeight = 180;
        comboBox.AccessibleDescription = placeholder;
    }

    private static void ConfigureTextBox(TextBox textBox)
    {
        textBox.Dock = DockStyle.Fill;
        textBox.BorderStyle = BorderStyle.FixedSingle;
        textBox.BackColor = AppTheme.Surface;
        textBox.ForeColor = AppTheme.TextPrimary;
        textBox.Margin = Padding.Empty;
    }

    private static TableLayoutPanel CreateFieldGrid(int rows)
    {
        var grid = new TableLayoutPanel();
        ConfigureFieldGrid(grid, rows);
        return grid;
    }

    private static void ConfigureFieldGrid(TableLayoutPanel grid, int rows)
    {
        grid.SuspendLayout();
        grid.Controls.Clear();
        grid.ColumnStyles.Clear();
        grid.RowStyles.Clear();
        grid.Dock = DockStyle.Top;
        grid.AutoSize = true;
        grid.ColumnCount = 4;
        grid.RowCount = rows;
        grid.Margin = Padding.Empty;
        grid.Padding = Padding.Empty;
        for (var index = 0; index < 4; index++)
        {
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
        }
        for (var index = 0; index < rows; index++)
        {
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 64F));
        }
        grid.ResumeLayout();
    }

    private static void PrepareDynamicField(Panel field, string label, Control input)
    {
        ConfigureFieldContainer(field);
        field.Controls.Add(CreateField(label, input));
    }

    private static void AddExistingField(
        TableLayoutPanel grid,
        Control field,
        int column,
        int row,
        int columnSpan = 1)
    {
        grid.Controls.Add(field, column, row);
        grid.SetColumnSpan(field, columnSpan);
    }

    private static Control CreateSection(string title, Control content)
    {
        var section = new Panel
        {
            Width = 780,
            Height = content.Height + 35,
            Margin = new Padding(0, 0, 0, 12),
            Padding = new Padding(0, 26, 0, 0)
        };
        section.Controls.Add(content);
        content.Dock = DockStyle.Fill;
        section.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
            ForeColor = AppTheme.TextPrimary,
            Location = new Point(0, 2)
        });
        return section;
    }

    private static void AddField(TableLayoutPanel grid, int column, int row, string label, Control input, int columnSpan = 1)
    {
        var field = CreateField(label, input);
        grid.Controls.Add(field, column, row);
        if (columnSpan > 1)
        {
            grid.SetColumnSpan(field, columnSpan);
        }
    }

    private static Panel CreateField(string label, Control input)
    {
        var field = new Panel();
        ConfigureFieldContainer(field);
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 21F));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 29F));

        var caption = new Label
        {
            Text = label,
            Dock = DockStyle.Fill,
            ForeColor = AppTheme.TextSecondary,
            Font = new Font("Segoe UI", 8.5F),
            TextAlign = ContentAlignment.MiddleLeft
        };
        input.Dock = DockStyle.Fill;
        layout.Controls.Add(caption, 0, 0);
        layout.Controls.Add(input, 0, 1);
        field.Controls.Add(layout);
        return field;
    }

    private static void ConfigureFieldContainer(Panel panel)
    {
        panel.Dock = DockStyle.Fill;
        panel.Margin = new Padding(0, 0, 12, 10);
        panel.Padding = Padding.Empty;
    }

    private static void AddSummaryRow(
        TableLayoutPanel layout,
        int row,
        string defaultLabel,
        Label value,
        Label? customLabel = null,
        bool emphasize = false)
    {
        var label = customLabel ?? new Label();
        label.Text = defaultLabel;
        label.AutoSize = true;
        label.ForeColor = AppTheme.TextSecondary;
        label.Font = new Font("Segoe UI", 8.5F, emphasize ? FontStyle.Bold : FontStyle.Regular);
        label.Margin = new Padding(0, 6, 0, 7);

        value.AutoSize = true;
        value.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        value.TextAlign = ContentAlignment.TopRight;
        value.ForeColor = emphasize ? AppTheme.Accent : AppTheme.TextPrimary;
        value.Font = new Font("Segoe UI", emphasize ? 11F : 8.5F, emphasize ? FontStyle.Bold : FontStyle.Regular);
        value.Margin = new Padding(0, 6, 0, 7);

        layout.Controls.Add(label, 0, row);
        layout.Controls.Add(value, 1, row);
    }
}
