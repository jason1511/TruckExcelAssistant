using TruckExcelAssistant.Models;
using TruckExcelAssistant.Services;

namespace TruckExcelAssistant;

public sealed class SettingsControl : UserControl
{
    private readonly DatabaseService _database;
    private readonly TextBox _companyName = new();
    private readonly TextBox _companyAddress = new();
    private readonly Panel _companyAddressFrame = new();
    private readonly TextBox _city = new();
    private readonly TextBox _bankName = new();
    private readonly TextBox _bankAccount = new();
    private readonly TextBox _bankHolder = new();
    private readonly TextBox _signerName = new();
    private readonly TextBox _invoicePrefix = new();
    private readonly NumericUpDown _sequenceDigits = new();
    private readonly ComboBox _defaultLayout = new();
    private readonly TextBox _exportDirectory = new();
    private readonly Label _numberPreview = new();
    private readonly Label _saveStatus = new();

    public SettingsControl(DatabaseService database)
    {
        _database = database;
        Dock = DockStyle.Fill;
        BackColor = AppTheme.WindowBackground;
        ForeColor = AppTheme.TextPrimary;
        Padding = new Padding(22, 18, 22, 22);
        AutoScaleMode = AutoScaleMode.Dpi;
        ConfigureInputs();
        BuildLayout();
        WireEvents();
        LoadSettings();
    }

    public event EventHandler? SettingsSaved;

    public void LoadSettings()
    {
        var settings = _database.GetSettings();
        _companyName.Text = settings.CompanyName;
        _companyAddress.Text = settings.CompanyAddress;
        _city.Text = settings.City;
        _bankName.Text = settings.BankName;
        _bankAccount.Text = settings.BankAccountNumber;
        _bankHolder.Text = settings.BankAccountHolder;
        _signerName.Text = settings.SignerName;
        _invoicePrefix.Text = settings.InvoicePrefix;
        _sequenceDigits.Value = settings.InvoiceSequenceDigits;
        _defaultLayout.SelectedIndex = settings.DefaultInvoiceLayout == OutputLayout.CompactInvoice ? 0 : 1;
        _exportDirectory.Text = settings.DefaultExportDirectory;
        _saveStatus.Text = string.Empty;
        UpdateNumberPreview();
    }

    private void ConfigureInputs()
    {
        foreach (var textBox in new[]
                 {
                     _companyName, _companyAddress, _city, _bankName, _bankAccount,
                     _bankHolder, _signerName, _invoicePrefix, _exportDirectory
                 })
        {
            textBox.BorderStyle = BorderStyle.FixedSingle;
        }
        _companyAddress.Multiline = true;
        _companyAddress.ScrollBars = ScrollBars.Vertical;
        _companyAddress.PlaceholderText = "Alamat perusahaan yang dicetak pada invoice";
        _companyAddress.BorderStyle = BorderStyle.None;
        _companyAddress.BackColor = AppTheme.Surface;
        _companyAddress.Margin = Padding.Empty;
        var addressInner = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.Surface,
            Padding = new Padding(6, 4, 3, 3),
            Margin = Padding.Empty
        };
        _companyAddress.Dock = DockStyle.Fill;
        addressInner.Controls.Add(_companyAddress);
        _companyAddressFrame.Dock = DockStyle.Fill;
        _companyAddressFrame.BackColor = AppTheme.InputBorder;
        _companyAddressFrame.Padding = new Padding(1);
        _companyAddressFrame.Margin = Padding.Empty;
        _companyAddressFrame.Controls.Add(addressInner);
        _companyAddress.Enter += (_, _) => _companyAddressFrame.BackColor = AppTheme.Accent;
        _companyAddress.Leave += (_, _) => _companyAddressFrame.BackColor = AppTheme.InputBorder;
        _invoicePrefix.CharacterCasing = CharacterCasing.Upper;
        _invoicePrefix.MaxLength = 8;
        _sequenceDigits.Minimum = 3;
        _sequenceDigits.Maximum = 6;
        _sequenceDigits.TextAlign = HorizontalAlignment.Right;
        _defaultLayout.DropDownStyle = ComboBoxStyle.DropDownList;
        _defaultLayout.Items.AddRange(["Invoice ringkas", "Invoice lengkap"]);
        _defaultLayout.SelectedIndex = 1;
        _exportDirectory.PlaceholderText = "Kosongkan untuk memakai folder Exports di samping aplikasi";
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
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 62F));

        var heading = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
        heading.Controls.Add(new Label
        {
            Text = "Pengaturan",
            AutoSize = true,
            Font = new Font("Segoe UI", 17F, FontStyle.Bold),
            ForeColor = AppTheme.TextPrimary,
            Location = new Point(0, 0)
        });
        heading.Controls.Add(new Label
        {
            Text = "Informasi ini disimpan lokal dan digunakan pada invoice Excel berikutnya.",
            AutoSize = true,
            Font = new Font("Segoe UI", 9F),
            ForeColor = AppTheme.TextSecondary,
            Location = new Point(2, 38)
        });

        var scroll = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = AppTheme.Surface,
            BorderStyle = BorderStyle.FixedSingle,
            Padding = new Padding(18),
            Margin = new Padding(0, 0, 0, 12)
        };
        var sections = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            RowCount = 3,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        sections.Controls.Add(BuildCompanySection(), 0, 0);
        sections.Controls.Add(BuildPaymentSection(), 0, 1);
        sections.Controls.Add(BuildInvoiceSection(), 0, 2);
        scroll.Controls.Add(sections);

        var footer = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = new Padding(0, 4, 0, 0)
        };
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190F));
        _saveStatus.Dock = DockStyle.Fill;
        _saveStatus.TextAlign = ContentAlignment.MiddleRight;
        _saveStatus.ForeColor = AppTheme.Accent;
        _saveStatus.Margin = new Padding(0, 0, 14, 0);
        var save = AppTheme.CreatePrimaryButton("Simpan pengaturan");
        save.Dock = DockStyle.Fill;
        save.Margin = Padding.Empty;
        save.Click += (_, _) => SaveSettings();
        footer.Controls.Add(_saveStatus, 0, 0);
        footer.Controls.Add(save, 1, 0);

        root.Controls.Add(heading, 0, 0);
        root.Controls.Add(scroll, 0, 1);
        root.Controls.Add(footer, 0, 2);
        Controls.Add(root);
    }

    private Control BuildCompanySection()
    {
        var grid = CreateSectionGrid(2, 2, 164);
        AddField(grid, 0, 0, "Nama perusahaan", _companyName);
        AddField(grid, 1, 0, "Kota penerbit invoice", _city);
        AddField(grid, 0, 1, "Alamat perusahaan", _companyAddressFrame, 2);
        return CreateSection("Identitas perusahaan", grid);
    }

    private Control BuildPaymentSection()
    {
        var grid = CreateSectionGrid(3, 1, 92);
        AddField(grid, 0, 0, "Nama bank", _bankName);
        AddField(grid, 1, 0, "Nomor rekening", _bankAccount);
        AddField(grid, 2, 0, "Nama pemilik rekening", _bankHolder);
        return CreateSection("Informasi pembayaran", grid);
    }

    private Control BuildInvoiceSection()
    {
        var grid = CreateSectionGrid(4, 3, 210);
        AddField(grid, 0, 0, "Nama penandatangan", _signerName, 2);
        AddField(grid, 2, 0, "Awalan nomor", _invoicePrefix);
        AddField(grid, 3, 0, "Jumlah digit urutan", _sequenceDigits);
        AddField(grid, 0, 1, "Layout invoice bawaan", _defaultLayout, 2);
        var previewPanel = new Panel { Dock = DockStyle.Fill, Margin = new Padding(7, 0, 0, 6) };
        previewPanel.Controls.Add(new Label
        {
            Text = "Pratinjau nomor berikutnya",
            AutoSize = true,
            ForeColor = AppTheme.TextSecondary,
            Font = new Font("Segoe UI", 8.5F),
            Location = new Point(0, 0)
        });
        _numberPreview.AutoSize = true;
        _numberPreview.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
        _numberPreview.ForeColor = AppTheme.Accent;
        _numberPreview.Location = new Point(0, 22);
        previewPanel.Controls.Add(_numberPreview);
        grid.Controls.Add(previewPanel, 2, 1);
        grid.SetColumnSpan(previewPanel, 2);

        var folderPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        folderPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        folderPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 125F));
        _exportDirectory.Dock = DockStyle.Fill;
        _exportDirectory.Margin = Padding.Empty;
        var browse = AppTheme.CreateSecondaryButton("Pilih folder");
        browse.Dock = DockStyle.Fill;
        browse.Margin = new Padding(8, 0, 0, 0);
        browse.Click += (_, _) => BrowseFolder();
        folderPanel.Controls.Add(_exportDirectory, 0, 0);
        folderPanel.Controls.Add(browse, 1, 0);
        AddField(grid, 0, 2, "Folder penyimpanan bawaan", folderPanel, 4);
        return CreateSection("Invoice dan penyimpanan", grid);
    }

    private void WireEvents()
    {
        _invoicePrefix.TextChanged += (_, _) => UpdateNumberPreview();
        _sequenceDigits.ValueChanged += (_, _) => UpdateNumberPreview();
    }

    private void BrowseFolder()
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "Pilih folder bawaan untuk menyimpan Excel",
            UseDescriptionForTitle = true,
            SelectedPath = Directory.Exists(_exportDirectory.Text) ? _exportDirectory.Text : string.Empty,
            ShowNewFolderButton = true
        };
        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _exportDirectory.Text = dialog.SelectedPath;
        }
    }

    private void SaveSettings()
    {
        var prefix = SanitizePrefix(_invoicePrefix.Text);
        if (string.IsNullOrWhiteSpace(prefix))
        {
            MessageBox.Show("Awalan nomor invoice harus berisi huruf atau angka.", "Pengaturan belum lengkap", MessageBoxButtons.OK, MessageBoxIcon.Information);
            _invoicePrefix.Focus();
            return;
        }
        if (string.IsNullOrWhiteSpace(_city.Text))
        {
            MessageBox.Show("Isi kota penerbit invoice.", "Pengaturan belum lengkap", MessageBoxButtons.OK, MessageBoxIcon.Information);
            _city.Focus();
            return;
        }

        var layout = _defaultLayout.SelectedIndex == 0
            ? OutputLayout.CompactInvoice
            : OutputLayout.CompleteInvoice;
        _database.SaveSettings(new AppSettings(
            _companyName.Text,
            _companyAddress.Text,
            _city.Text,
            _bankName.Text,
            _bankAccount.Text,
            _bankHolder.Text,
            _signerName.Text,
            prefix,
            Decimal.ToInt32(_sequenceDigits.Value),
            layout,
            _exportDirectory.Text));
        _invoicePrefix.Text = prefix;
        _saveStatus.Text = $"Tersimpan {DateTime.Now:HH:mm}";
        SettingsSaved?.Invoke(this, EventArgs.Empty);
        UpdateNumberPreview();
    }

    private void UpdateNumberPreview()
    {
        var prefix = SanitizePrefix(_invoicePrefix.Text);
        var digits = Decimal.ToInt32(_sequenceDigits.Value);
        _numberPreview.Text = $"{(string.IsNullOrWhiteSpace(prefix) ? "TJ" : prefix)}-{DateTime.Today:yyyyMMdd}-{1.ToString($"D{digits}")}";
    }

    private static string SanitizePrefix(string value) => new(value
        .Trim()
        .ToUpperInvariant()
        .Where(char.IsLetterOrDigit)
        .ToArray());

    private static TableLayoutPanel CreateSectionGrid(int columns, int rows, int height)
    {
        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = height,
            ColumnCount = columns,
            RowCount = rows,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        for (var index = 0; index < columns; index++)
        {
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F / columns));
        }
        for (var index = 0; index < rows; index++)
        {
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100F / rows));
        }
        return grid;
    }

    private static Control CreateSection(string title, Control content)
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            RowCount = 2,
            Margin = new Padding(0, 0, 0, 18),
            Padding = Padding.Empty
        };
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.Controls.Add(new Label
        {
            Text = title,
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            ForeColor = AppTheme.TextPrimary,
            TextAlign = ContentAlignment.MiddleLeft
        }, 0, 0);
        panel.Controls.Add(content, 0, 1);
        return panel;
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
