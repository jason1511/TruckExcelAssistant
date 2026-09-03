using System.Globalization;
using Microsoft.Data.Sqlite;
using TruckExcelAssistant.Models;

namespace TruckExcelAssistant.Services;

public sealed class DatabaseService
{
    private readonly string _connectionString;

    public DatabaseService(string? databasePath = null)
    {
        DatabasePath = databasePath ?? Path.Combine(AppContext.BaseDirectory, "truck_excel_assistant.db");
        _connectionString = new SqliteConnectionStringBuilder
        {
            DataSource = DatabasePath,
            Mode = SqliteOpenMode.ReadWriteCreate
        }.ToString();
    }

    public string DatabasePath { get; }

    public void Initialize()
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = """
            CREATE TABLE IF NOT EXISTS hauls (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                haul_date TEXT NOT NULL,
                licence_plate TEXT NOT NULL DEFAULT '',
                cargo TEXT NOT NULL DEFAULT '',
                customer TEXT NOT NULL DEFAULT '',
                origin TEXT NOT NULL DEFAULT '',
                destination TEXT NOT NULL DEFAULT '',
                loaded_weight_kg REAL NOT NULL DEFAULT 0,
                received_weight_kg REAL NOT NULL DEFAULT 0,
                rate_per_kg REAL NOT NULL DEFAULT 0,
                bon_sangu REAL NOT NULL DEFAULT 0,
                rejection_cost REAL NOT NULL DEFAULT 0,
                claim_amount REAL NOT NULL DEFAULT 0,
                driver_road_money REAL NOT NULL DEFAULT 0,
                other_expense REAL NOT NULL DEFAULT 0,
                notes TEXT NOT NULL DEFAULT '',
                preview_layout INTEGER NOT NULL DEFAULT 2,
                status TEXT NOT NULL DEFAULT 'Saved' CHECK (status IN ('Draft', 'Saved')),
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                deleted_at TEXT NULL
            );

            CREATE INDEX IF NOT EXISTS ix_hauls_date ON hauls(haul_date DESC);
            CREATE INDEX IF NOT EXISTS ix_hauls_plate ON hauls(licence_plate COLLATE NOCASE);
            CREATE INDEX IF NOT EXISTS ix_hauls_customer ON hauls(customer COLLATE NOCASE);

            CREATE TABLE IF NOT EXISTS invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_number TEXT NOT NULL UNIQUE,
                invoice_date TEXT NOT NULL,
                customer TEXT NOT NULL,
                layout INTEGER NOT NULL,
                total_amount REAL NOT NULL DEFAULT 0,
                file_path TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'Generated' CHECK (status IN ('Generated', 'Paid', 'Cancelled')),
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS invoice_hauls (
                invoice_id INTEGER NOT NULL,
                haul_id INTEGER NOT NULL,
                PRIMARY KEY (invoice_id, haul_id),
                FOREIGN KEY (invoice_id) REFERENCES invoices(id) ON DELETE CASCADE,
                FOREIGN KEY (haul_id) REFERENCES hauls(id)
            );

            CREATE INDEX IF NOT EXISTS ix_invoices_date ON invoices(invoice_date DESC);
            CREATE INDEX IF NOT EXISTS ix_invoices_customer ON invoices(customer COLLATE NOCASE);

            CREATE TABLE IF NOT EXISTS app_settings (
                setting_key TEXT PRIMARY KEY,
                setting_value TEXT NOT NULL
            );
            """;
        command.ExecuteNonQuery();

        if (!ColumnExists(connection, "hauls", "deleted_at"))
        {
            using var migration = connection.CreateCommand();
            migration.CommandText = "ALTER TABLE hauls ADD COLUMN deleted_at TEXT NULL;";
            migration.ExecuteNonQuery();
        }

        using var version = connection.CreateCommand();
        version.CommandText = "PRAGMA user_version = 4;";
        version.ExecuteNonQuery();
    }

    public long AddHaul(HaulDraft draft, HaulStatus status)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = """
            INSERT INTO hauls (
                haul_date, licence_plate, cargo, customer, origin, destination,
                loaded_weight_kg, received_weight_kg, rate_per_kg, bon_sangu,
                rejection_cost, claim_amount, driver_road_money, other_expense,
                notes, preview_layout, status, created_at, updated_at
            ) VALUES (
                $date, $plate, $cargo, $customer, $origin, $destination,
                $loadedWeight, $receivedWeight, $rate, $bonSangu,
                $rejectionCost, $claimAmount, $driverRoadMoney, $otherExpense,
                $notes, $layout, $status, $createdAt, $updatedAt
            );
            SELECT last_insert_rowid();
            """;

        var now = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
        command.Parameters.AddWithValue("$date", draft.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        command.Parameters.AddWithValue("$plate", draft.LicencePlate);
        command.Parameters.AddWithValue("$cargo", draft.Cargo);
        command.Parameters.AddWithValue("$customer", draft.Customer);
        command.Parameters.AddWithValue("$origin", draft.Origin);
        command.Parameters.AddWithValue("$destination", draft.Destination);
        command.Parameters.AddWithValue("$loadedWeight", Convert.ToDouble(draft.LoadedWeightKg));
        command.Parameters.AddWithValue("$receivedWeight", Convert.ToDouble(draft.ReceivedWeightKg));
        command.Parameters.AddWithValue("$rate", Convert.ToDouble(draft.RatePerKg));
        command.Parameters.AddWithValue("$bonSangu", Convert.ToDouble(draft.BonSangu));
        command.Parameters.AddWithValue("$rejectionCost", Convert.ToDouble(draft.RejectionCost));
        command.Parameters.AddWithValue("$claimAmount", Convert.ToDouble(draft.ClaimAmount));
        command.Parameters.AddWithValue("$driverRoadMoney", Convert.ToDouble(draft.DriverRoadMoney));
        command.Parameters.AddWithValue("$otherExpense", Convert.ToDouble(draft.OtherExpense));
        command.Parameters.AddWithValue("$notes", draft.Notes);
        command.Parameters.AddWithValue("$layout", (int)draft.Layout);
        command.Parameters.AddWithValue("$status", status.ToString());
        command.Parameters.AddWithValue("$createdAt", now);
        command.Parameters.AddWithValue("$updatedAt", now);

        return Convert.ToInt64(command.ExecuteScalar(), CultureInfo.InvariantCulture);
    }

    public void UpdateHaul(long id, HaulDraft draft, HaulStatus status)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = """
            UPDATE hauls SET
                haul_date = $date,
                licence_plate = $plate,
                cargo = $cargo,
                customer = $customer,
                origin = $origin,
                destination = $destination,
                loaded_weight_kg = $loadedWeight,
                received_weight_kg = $receivedWeight,
                rate_per_kg = $rate,
                bon_sangu = $bonSangu,
                rejection_cost = $rejectionCost,
                claim_amount = $claimAmount,
                driver_road_money = $driverRoadMoney,
                other_expense = $otherExpense,
                notes = $notes,
                preview_layout = $layout,
                status = $status,
                updated_at = $updatedAt
            WHERE id = $id AND deleted_at IS NULL;
            """;
        AddDraftParameters(command, draft, status);
        command.Parameters.AddWithValue("$id", id);
        if (command.ExecuteNonQuery() == 0)
        {
            throw new InvalidOperationException("Data angkutan tidak ditemukan atau sudah berada di Sampah.");
        }
    }

    public HaulRecord? GetHaul(long id)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = $"""
            {SelectHaulSql}
            WHERE id = $id;
            """;
        command.Parameters.AddWithValue("$id", id);
        using var reader = command.ExecuteReader();
        return reader.Read() ? ReadHaul(reader) : null;
    }

    public IReadOnlyList<HaulRecord> GetHauls(
        string? searchText = null,
        HaulStatus? status = null,
        bool deletedOnly = false)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        var search = searchText?.Trim() ?? string.Empty;
        command.CommandText = $"""
            {SelectHaulSql}
            WHERE (
                    $search = ''
                 OR licence_plate LIKE $pattern COLLATE NOCASE
                 OR cargo LIKE $pattern COLLATE NOCASE
                 OR customer LIKE $pattern COLLATE NOCASE
                 OR origin LIKE $pattern COLLATE NOCASE
                 OR destination LIKE $pattern COLLATE NOCASE
                 OR notes LIKE $pattern COLLATE NOCASE
            )
              AND (($deletedOnly = 1 AND deleted_at IS NOT NULL)
                OR ($deletedOnly = 0 AND deleted_at IS NULL))
              AND ($status = '' OR status = $status)
            ORDER BY haul_date DESC, id DESC
            LIMIT 500;
            """;
        command.Parameters.AddWithValue("$search", search);
        command.Parameters.AddWithValue("$pattern", $"%{search}%");
        command.Parameters.AddWithValue("$deletedOnly", deletedOnly ? 1 : 0);
        command.Parameters.AddWithValue("$status", status?.ToString() ?? string.Empty);

        var records = new List<HaulRecord>();
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            records.Add(ReadHaul(reader));
        }
        return records;
    }

    public IReadOnlyList<HaulRecord> GetSavedHaulsForExport(
        DateTime from,
        DateTime to,
        bool excludeAlreadyInvoiced = false)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = $"""
            {SelectHaulSql}
            WHERE deleted_at IS NULL
              AND status = 'Saved'
              AND haul_date >= $from
              AND haul_date <= $to
              AND ($excludeAlreadyInvoiced = 0 OR NOT EXISTS (
                    SELECT 1
                    FROM invoice_hauls ih
                    INNER JOIN invoices i ON i.id = ih.invoice_id
                    WHERE ih.haul_id = hauls.id AND i.status <> 'Cancelled'
              ))
            ORDER BY haul_date, id;
            """;
        command.Parameters.AddWithValue("$from", from.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        command.Parameters.AddWithValue("$to", to.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        command.Parameters.AddWithValue("$excludeAlreadyInvoiced", excludeAlreadyInvoiced ? 1 : 0);

        var records = new List<HaulRecord>();
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            records.Add(ReadHaul(reader));
        }
        return records;
    }

    public string GetNextInvoiceNumber(DateTime invoiceDate)
    {
        var settings = GetSettings();
        var prefix = $"{settings.InvoicePrefix}-{invoiceDate:yyyyMMdd}-";
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT COALESCE(MAX(CAST(SUBSTR(invoice_number, LENGTH($prefix) + 1) AS INTEGER)), 0) + 1
            FROM invoices
            WHERE invoice_number LIKE $pattern;
            """;
        command.Parameters.AddWithValue("$prefix", prefix);
        command.Parameters.AddWithValue("$pattern", $"{prefix}%");
        var sequence = Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture);
        return $"{prefix}{sequence.ToString($"D{settings.InvoiceSequenceDigits}", CultureInfo.InvariantCulture)}";
    }

    public AppSettings GetSettings()
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT setting_key, setting_value FROM app_settings;";
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            values[reader.GetString(0)] = reader.GetString(1);
        }

        var defaults = AppSettings.Default;
        var layout = values.TryGetValue("default_invoice_layout", out var layoutValue)
                     && Enum.TryParse<OutputLayout>(layoutValue, out var parsedLayout)
                     && parsedLayout is OutputLayout.CompactInvoice or OutputLayout.CompleteInvoice
            ? parsedLayout
            : defaults.DefaultInvoiceLayout;
        var digits = values.TryGetValue("invoice_sequence_digits", out var digitsValue)
                     && int.TryParse(digitsValue, CultureInfo.InvariantCulture, out var parsedDigits)
            ? Math.Clamp(parsedDigits, 3, 6)
            : defaults.InvoiceSequenceDigits;

        return new AppSettings(
            GetSetting(values, "company_name", defaults.CompanyName),
            GetSetting(values, "company_address", defaults.CompanyAddress),
            GetSetting(values, "city", defaults.City),
            GetSetting(values, "bank_name", defaults.BankName),
            GetSetting(values, "bank_account_number", defaults.BankAccountNumber),
            GetSetting(values, "bank_account_holder", defaults.BankAccountHolder),
            GetSetting(values, "signer_name", defaults.SignerName),
            NormalizeInvoicePrefix(GetSetting(values, "invoice_prefix", defaults.InvoicePrefix)),
            digits,
            layout,
            GetSetting(values, "default_export_directory", defaults.DefaultExportDirectory));
    }

    public void SaveSettings(AppSettings settings)
    {
        using var connection = OpenConnection();
        using var transaction = connection.BeginTransaction();
        var values = new Dictionary<string, string>
        {
            ["company_name"] = settings.CompanyName.Trim(),
            ["company_address"] = settings.CompanyAddress.Trim(),
            ["city"] = settings.City.Trim(),
            ["bank_name"] = settings.BankName.Trim(),
            ["bank_account_number"] = settings.BankAccountNumber.Trim(),
            ["bank_account_holder"] = settings.BankAccountHolder.Trim(),
            ["signer_name"] = settings.SignerName.Trim(),
            ["invoice_prefix"] = NormalizeInvoicePrefix(settings.InvoicePrefix),
            ["invoice_sequence_digits"] = Math.Clamp(settings.InvoiceSequenceDigits, 3, 6).ToString(CultureInfo.InvariantCulture),
            ["default_invoice_layout"] = settings.DefaultInvoiceLayout.ToString(),
            ["default_export_directory"] = settings.DefaultExportDirectory.Trim()
        };
        foreach (var pair in values)
        {
            using var command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText = """
                INSERT INTO app_settings (setting_key, setting_value)
                VALUES ($key, $value)
                ON CONFLICT(setting_key) DO UPDATE SET setting_value = excluded.setting_value;
                """;
            command.Parameters.AddWithValue("$key", pair.Key);
            command.Parameters.AddWithValue("$value", pair.Value);
            command.ExecuteNonQuery();
        }
        transaction.Commit();
    }

    public void RecordGeneratedInvoice(
        string invoiceNumber,
        DateTime invoiceDate,
        string customer,
        OutputLayout layout,
        decimal totalAmount,
        string filePath,
        IReadOnlyCollection<long> haulIds)
    {
        using var connection = OpenConnection();
        using var transaction = connection.BeginTransaction();
        using var invoice = connection.CreateCommand();
        invoice.Transaction = transaction;
        invoice.CommandText = """
            INSERT INTO invoices (
                invoice_number, invoice_date, customer, layout, total_amount,
                file_path, status, created_at
            ) VALUES (
                $number, $date, $customer, $layout, $total,
                $path, 'Generated', $createdAt
            );
            SELECT last_insert_rowid();
            """;
        invoice.Parameters.AddWithValue("$number", invoiceNumber);
        invoice.Parameters.AddWithValue("$date", invoiceDate.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        invoice.Parameters.AddWithValue("$customer", customer.Trim());
        invoice.Parameters.AddWithValue("$layout", (int)layout);
        invoice.Parameters.AddWithValue("$total", Convert.ToDouble(totalAmount));
        invoice.Parameters.AddWithValue("$path", filePath);
        invoice.Parameters.AddWithValue("$createdAt", DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
        var invoiceId = Convert.ToInt64(invoice.ExecuteScalar(), CultureInfo.InvariantCulture);

        foreach (var haulId in haulIds)
        {
            using var link = connection.CreateCommand();
            link.Transaction = transaction;
            link.CommandText = "INSERT INTO invoice_hauls (invoice_id, haul_id) VALUES ($invoiceId, $haulId);";
            link.Parameters.AddWithValue("$invoiceId", invoiceId);
            link.Parameters.AddWithValue("$haulId", haulId);
            link.ExecuteNonQuery();
        }
        transaction.Commit();
    }

    public IReadOnlyList<InvoiceRecord> GetInvoices(string? searchText = null, InvoiceStatus? status = null)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        var search = searchText?.Trim() ?? string.Empty;
        command.CommandText = """
            SELECT i.id, i.invoice_number, i.invoice_date, i.customer, i.layout,
                   i.total_amount, i.file_path, i.status, i.created_at,
                   COUNT(ih.haul_id) AS haul_count
            FROM invoices i
            LEFT JOIN invoice_hauls ih ON ih.invoice_id = i.id
            WHERE ($search = ''
                   OR i.invoice_number LIKE $pattern COLLATE NOCASE
                   OR i.customer LIKE $pattern COLLATE NOCASE
                   OR i.file_path LIKE $pattern COLLATE NOCASE)
              AND ($status = '' OR i.status = $status)
            GROUP BY i.id
            ORDER BY i.invoice_date DESC, i.id DESC
            LIMIT 500;
            """;
        command.Parameters.AddWithValue("$search", search);
        command.Parameters.AddWithValue("$pattern", $"%{search}%");
        command.Parameters.AddWithValue("$status", status?.ToString() ?? string.Empty);

        var invoices = new List<InvoiceRecord>();
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            var layoutValue = reader.GetInt32(4);
            var layout = Enum.IsDefined(typeof(OutputLayout), layoutValue)
                ? (OutputLayout)layoutValue
                : OutputLayout.CompleteInvoice;
            var invoiceStatus = Enum.TryParse<InvoiceStatus>(reader.GetString(7), out var parsed)
                ? parsed
                : InvoiceStatus.Generated;
            invoices.Add(new InvoiceRecord(
                reader.GetInt64(0),
                reader.GetString(1),
                DateTime.ParseExact(reader.GetString(2), "yyyy-MM-dd", CultureInfo.InvariantCulture),
                reader.GetString(3),
                layout,
                ToDecimal(reader, 5),
                reader.GetString(6),
                invoiceStatus,
                DateTime.Parse(reader.GetString(8), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind),
                reader.GetInt32(9)));
        }
        return invoices;
    }

    public IReadOnlyList<HaulRecord> GetInvoiceHauls(long invoiceId)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT h.id, h.haul_date, h.licence_plate, h.cargo, h.customer, h.origin, h.destination,
                   h.loaded_weight_kg, h.received_weight_kg, h.rate_per_kg, h.bon_sangu,
                   h.rejection_cost, h.claim_amount, h.driver_road_money, h.other_expense,
                   h.notes, h.preview_layout, h.status, h.created_at, h.updated_at, h.deleted_at
            FROM hauls h
            INNER JOIN invoice_hauls ih ON ih.haul_id = h.id
            WHERE ih.invoice_id = $invoiceId
            ORDER BY h.haul_date, h.id;
            """;
        command.Parameters.AddWithValue("$invoiceId", invoiceId);
        var records = new List<HaulRecord>();
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            records.Add(ReadHaul(reader));
        }
        return records;
    }

    public void UpdateInvoiceStatus(long invoiceId, InvoiceStatus status)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = "UPDATE invoices SET status = $status WHERE id = $id;";
        command.Parameters.AddWithValue("$status", status.ToString());
        command.Parameters.AddWithValue("$id", invoiceId);
        if (command.ExecuteNonQuery() == 0)
        {
            throw new InvalidOperationException("Invoice tidak ditemukan.");
        }
    }

    public void UpdateInvoiceFilePath(long invoiceId, string filePath)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = "UPDATE invoices SET file_path = $path WHERE id = $id;";
        command.Parameters.AddWithValue("$path", filePath);
        command.Parameters.AddWithValue("$id", invoiceId);
        if (command.ExecuteNonQuery() == 0)
        {
            throw new InvalidOperationException("Invoice tidak ditemukan.");
        }
    }

    public void MoveToTrash(long id)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = """
            UPDATE hauls
            SET deleted_at = $deletedAt, updated_at = $deletedAt
            WHERE id = $id AND deleted_at IS NULL;
            """;
        command.Parameters.AddWithValue("$id", id);
        command.Parameters.AddWithValue("$deletedAt", DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
        command.ExecuteNonQuery();
    }

    public void RestoreFromTrash(long id)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = """
            UPDATE hauls
            SET deleted_at = NULL, updated_at = $updatedAt
            WHERE id = $id AND deleted_at IS NOT NULL;
            """;
        command.Parameters.AddWithValue("$id", id);
        command.Parameters.AddWithValue("$updatedAt", DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
        command.ExecuteNonQuery();
    }

    public IReadOnlyList<string> GetSuggestions(SuggestionField field)
    {
        var column = field switch
        {
            SuggestionField.LicencePlate => "licence_plate",
            SuggestionField.Cargo => "cargo",
            SuggestionField.Customer => "customer",
            SuggestionField.Origin => "origin",
            SuggestionField.Destination => "destination",
            _ => throw new ArgumentOutOfRangeException(nameof(field), field, null)
        };

        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = $"""
            SELECT {column}
            FROM hauls
            WHERE TRIM({column}) <> '' AND deleted_at IS NULL
            GROUP BY {column} COLLATE NOCASE
            ORDER BY MAX(id) DESC
            LIMIT 100;
            """;

        var values = new List<string>();
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            values.Add(reader.GetString(0));
        }
        return values;
    }

    public int CountHauls()
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT COUNT(*) FROM hauls WHERE deleted_at IS NULL;";
        return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture);
    }

    private SqliteConnection OpenConnection()
    {
        var connection = new SqliteConnection(_connectionString);
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = "PRAGMA foreign_keys = ON; PRAGMA busy_timeout = 5000;";
        command.ExecuteNonQuery();
        return connection;
    }

    private static void AddDraftParameters(SqliteCommand command, HaulDraft draft, HaulStatus status)
    {
        command.Parameters.AddWithValue("$date", draft.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        command.Parameters.AddWithValue("$plate", draft.LicencePlate);
        command.Parameters.AddWithValue("$cargo", draft.Cargo);
        command.Parameters.AddWithValue("$customer", draft.Customer);
        command.Parameters.AddWithValue("$origin", draft.Origin);
        command.Parameters.AddWithValue("$destination", draft.Destination);
        command.Parameters.AddWithValue("$loadedWeight", Convert.ToDouble(draft.LoadedWeightKg));
        command.Parameters.AddWithValue("$receivedWeight", Convert.ToDouble(draft.ReceivedWeightKg));
        command.Parameters.AddWithValue("$rate", Convert.ToDouble(draft.RatePerKg));
        command.Parameters.AddWithValue("$bonSangu", Convert.ToDouble(draft.BonSangu));
        command.Parameters.AddWithValue("$rejectionCost", Convert.ToDouble(draft.RejectionCost));
        command.Parameters.AddWithValue("$claimAmount", Convert.ToDouble(draft.ClaimAmount));
        command.Parameters.AddWithValue("$driverRoadMoney", Convert.ToDouble(draft.DriverRoadMoney));
        command.Parameters.AddWithValue("$otherExpense", Convert.ToDouble(draft.OtherExpense));
        command.Parameters.AddWithValue("$notes", draft.Notes);
        command.Parameters.AddWithValue("$layout", (int)draft.Layout);
        command.Parameters.AddWithValue("$status", status.ToString());
        command.Parameters.AddWithValue("$updatedAt", DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
    }

    private static bool ColumnExists(SqliteConnection connection, string tableName, string columnName)
    {
        using var command = connection.CreateCommand();
        command.CommandText = $"PRAGMA table_info({tableName});";
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            if (reader.GetString(1).Equals(columnName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }
        return false;
    }

    private const string SelectHaulSql = """
        SELECT id, haul_date, licence_plate, cargo, customer, origin, destination,
               loaded_weight_kg, received_weight_kg, rate_per_kg, bon_sangu,
               rejection_cost, claim_amount, driver_road_money, other_expense,
               notes, preview_layout, status, created_at, updated_at, deleted_at
        FROM hauls
        """;

    private static HaulRecord ReadHaul(SqliteDataReader reader)
    {
        var layoutValue = reader.GetInt32(16);
        var layout = Enum.IsDefined(typeof(OutputLayout), layoutValue)
            ? (OutputLayout)layoutValue
            : OutputLayout.CompleteInvoice;
        var status = Enum.TryParse<HaulStatus>(reader.GetString(17), out var parsedStatus)
            ? parsedStatus
            : HaulStatus.Saved;

        var draft = new HaulDraft(
            DateTime.ParseExact(reader.GetString(1), "yyyy-MM-dd", CultureInfo.InvariantCulture),
            reader.GetString(2),
            reader.GetString(3),
            reader.GetString(4),
            reader.GetString(5),
            reader.GetString(6),
            ToDecimal(reader, 7),
            ToDecimal(reader, 8),
            ToDecimal(reader, 9),
            ToDecimal(reader, 10),
            ToDecimal(reader, 11),
            ToDecimal(reader, 12),
            ToDecimal(reader, 13),
            ToDecimal(reader, 14),
            reader.GetString(15),
            layout);

        return new HaulRecord(
            reader.GetInt64(0),
            draft,
            status,
            DateTime.Parse(reader.GetString(18), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind),
            DateTime.Parse(reader.GetString(19), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind),
            reader.IsDBNull(20)
                ? null
                : DateTime.Parse(reader.GetString(20), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind));
    }

    private static decimal ToDecimal(SqliteDataReader reader, int ordinal) =>
        (decimal)reader.GetDouble(ordinal);

    private static string GetSetting(
        IReadOnlyDictionary<string, string> settings,
        string key,
        string fallback) => settings.TryGetValue(key, out var value) ? value : fallback;

    private static string NormalizeInvoicePrefix(string value)
    {
        var normalized = new string(value
            .Trim()
            .ToUpperInvariant()
            .Where(char.IsLetterOrDigit)
            .ToArray());
        return string.IsNullOrWhiteSpace(normalized) ? "TJ" : normalized;
    }
}
