using System.Globalization;
using Microsoft.Data.Sqlite;
using TruckExcelAssistant.Models;

namespace TruckExcelAssistant.Services;

public sealed class DatabaseService
{
    private readonly string _connectionString;

    public DatabaseService()
    {
        DatabasePath = Path.Combine(AppContext.BaseDirectory, "truck_excel_assistant.db");
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
                updated_at TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS ix_hauls_date ON hauls(haul_date DESC);
            CREATE INDEX IF NOT EXISTS ix_hauls_plate ON hauls(licence_plate COLLATE NOCASE);
            CREATE INDEX IF NOT EXISTS ix_hauls_customer ON hauls(customer COLLATE NOCASE);
            PRAGMA user_version = 1;
            """;
        command.ExecuteNonQuery();
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

    public IReadOnlyList<HaulRecord> GetHauls(string? searchText = null)
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        var search = searchText?.Trim() ?? string.Empty;
        command.CommandText = """
            SELECT id, haul_date, licence_plate, cargo, customer, origin, destination,
                   loaded_weight_kg, received_weight_kg, rate_per_kg, bon_sangu,
                   rejection_cost, claim_amount, driver_road_money, other_expense,
                   notes, preview_layout, status, created_at, updated_at
            FROM hauls
            WHERE $search = ''
               OR licence_plate LIKE $pattern COLLATE NOCASE
               OR cargo LIKE $pattern COLLATE NOCASE
               OR customer LIKE $pattern COLLATE NOCASE
               OR origin LIKE $pattern COLLATE NOCASE
               OR destination LIKE $pattern COLLATE NOCASE
               OR notes LIKE $pattern COLLATE NOCASE
            ORDER BY haul_date DESC, id DESC
            LIMIT 500;
            """;
        command.Parameters.AddWithValue("$search", search);
        command.Parameters.AddWithValue("$pattern", $"%{search}%");

        var records = new List<HaulRecord>();
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            records.Add(ReadHaul(reader));
        }
        return records;
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
            WHERE TRIM({column}) <> ''
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
        command.CommandText = "SELECT COUNT(*) FROM hauls;";
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
            DateTime.Parse(reader.GetString(19), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind));
    }

    private static decimal ToDecimal(SqliteDataReader reader, int ordinal) =>
        Convert.ToDecimal(reader.GetDouble(ordinal), CultureInfo.InvariantCulture);
}
