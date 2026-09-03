using TruckExcelAssistant.Models;

namespace TruckExcelAssistant.Services;

public sealed class DemoDataSeeder
{
    private const string DemoTag = "[DATA CONTOH]";
    private readonly DatabaseService _database;
    private readonly ExcelExportService _exporter;

    public DemoDataSeeder(DatabaseService database, ExcelExportService exporter)
    {
        _database = database;
        _exporter = exporter;
    }

    public void Seed()
    {
        if (_database.HasDemoData())
        {
            throw new InvalidOperationException("Data contoh sudah ada. Hapus data contoh lama sebelum mengisinya kembali.");
        }
        try
        {
            SeedCore();
        }
        catch
        {
            Remove();
            throw;
        }
    }

    public void Remove()
    {
        var paths = _database.GetDemoInvoicePaths();
        _database.RemoveDemoData();
        foreach (var path in paths.Where(File.Exists))
        {
            File.Delete(path);
        }
    }

    private void SeedCore()
    {
        var today = DateTime.Today;
        var monthStart = new DateTime(today.Year, today.Month, 1);
        DateTime Day(int day) => monthStart.AddDays(Math.Min(day - 1, Math.Max(0, today.Day - 1)));

        var drafts = new[]
        {
            Draft(Day(1), "N 1001 XX", "Jagung", "PT Pangan Contoh", "Surabaya", "Semarang", 45_600, 45_350, 315, 2_900_000, 250_000, OutputLayout.CompleteInvoice),
            Draft(Day(2), "N 1002 YY", "Jagung", "PT Pangan Contoh", "Surabaya", "Solo", 46_100, 45_920, 315, 2_700_000, 180_000, OutputLayout.CompleteInvoice),
            Draft(Day(3), "N 1003 ZZ", "Jagung", "PT Pangan Contoh", "Gresik", "Kediri", 44_850, 44_700, 285, 2_500_000, 150_000, OutputLayout.CompleteInvoice, rejection: 200_000),
            Draft(Day(4), "P 2001 AA", "Jagung", "PT Pangan Contoh", "Jember", "Cirebon", 47_200, 46_650, 300, 3_200_000, 300_000, OutputLayout.CompleteInvoice, claim: 650_000),
            Draft(Day(5), "N 1001 XX", "SBM", "CV Logistik Demo", "Teluk Lamong", "Semarang", 44_300, 44_050, 145, 3_050_000, 125_000, OutputLayout.CompactInvoice, bon: 300_000),
            Draft(Day(6), "N 1002 YY", "SBM", "CV Logistik Demo", "Teluk Lamong", "Grobogan", 45_100, 44_900, 145, 3_100_000, 175_000, OutputLayout.CompactInvoice),
            Draft(Day(7), "N 1003 ZZ", "Tepung", "CV Logistik Demo", "Jakarta", "Pasuruan", 40_200, 40_000, 275, 5_500_000, 200_000, OutputLayout.CompactInvoice, bon: 250_000),
            Draft(Day(8), "P 2001 AA", "Pasir", "CV Logistik Demo", "Lumajang", "Balaraja", 45_250, 45_000, 310, 6_700_000, 225_000, OutputLayout.CompactInvoice),
            Draft(Day(9), "N 1001 XX", "Pupuk", "PT Pakan Uji", "Gresik", "Jatiroto", 45_100, 45_000, 90, 2_200_000, 100_000, OutputLayout.TruckLedger),
            Draft(Day(10), "N 1002 YY", "Dedak", "PT Pakan Uji", "Probolinggo", "Madiun", 26_400, 26_150, 250, 1_900_000, 150_000, OutputLayout.TruckLedger),
            Draft(Day(11), "N 1003 ZZ", "Jagung", "PT Pakan Uji", "Jember", "Surabaya", 46_300, 46_100, 95, 2_000_000, 125_000, OutputLayout.TruckLedger),
            Draft(Day(12), "P 2001 AA", "SBM", "PT Pakan Uji", "Teluk Lamong", "Batang", 40_100, 39_880, 165, 3_300_000, 175_000, OutputLayout.TruckLedger)
        };

        var records = new List<HaulRecord>();
        foreach (var draft in drafts)
        {
            var id = _database.AddHaul(draft, HaulStatus.Saved);
            records.Add(_database.GetHaul(id) ?? throw new InvalidOperationException("Data contoh gagal dibaca kembali."));
        }

        AddExpense(Day(2), "N 1001 XX", "Bahan bakar", "Solar perjalanan", 1_250_000);
        AddExpense(Day(3), "N 1002 YY", "Tol & parkir", "Tol dan parkir", 475_000);
        AddExpense(Day(5), "N 1003 ZZ", "Ban", "Ganti ban belakang", 1_850_000);
        AddExpense(Day(7), "P 2001 AA", "Servis & suku cadang", "Servis radiator", 925_000);
        AddExpense(Day(9), "N 1001 XX", "Pajak & administrasi", "Administrasi kendaraan", 350_000);
        AddExpense(Day(11), "N 1002 YY", "Lainnya", "Cuci kendaraan", 125_000);

        var settings = _database.GetSettings();
        var baseDirectory = !string.IsNullOrWhiteSpace(settings.DefaultExportDirectory)
                            && Directory.Exists(settings.DefaultExportDirectory)
            ? settings.DefaultExportDirectory
            : _exporter.ExportDirectory;
        var directory = Path.Combine(baseDirectory, "DataContoh");
        Directory.CreateDirectory(directory);
        CreateInvoice(records, "PT Pangan Contoh", today, directory, settings, OutputLayout.CompleteInvoice, markPaid: true);
        CreateInvoice(records, "CV Logistik Demo", today, directory, settings, OutputLayout.CompactInvoice, markPaid: false);
    }

    private void AddExpense(DateTime date, string plate, string category, string description, decimal amount) =>
        _database.AddExpense(date, plate, category, $"{DemoTag} {description}", amount);

    private void CreateInvoice(
        IReadOnlyList<HaulRecord> records,
        string customer,
        DateTime date,
        string directory,
        AppSettings settings,
        OutputLayout layout,
        bool markPaid)
    {
        var hauls = records.Where(item => item.Draft.Customer == customer).ToList();
        var number = _database.GetNextInvoiceNumber(date);
        var path = Path.Combine(directory, $"Invoice-{number}.xlsx");
        if (layout == OutputLayout.CompactInvoice)
        {
            _exporter.ExportCompactInvoice(hauls, customer, number, date, path, settings);
        }
        else
        {
            _exporter.ExportCompleteInvoice(hauls, customer, number, date, path, settings);
        }
        _database.RecordGeneratedInvoice(
            number, date, customer, layout,
            hauls.Sum(item => item.Draft.FinalAmount),
            path,
            hauls.Select(item => item.Id).ToList());
        if (markPaid)
        {
            var invoice = _database.GetInvoices(number).Single(item => item.InvoiceNumber == number);
            _database.UpdateInvoiceStatus(invoice.Id, InvoiceStatus.Paid);
        }
    }

    private static HaulDraft Draft(
        DateTime date,
        string plate,
        string cargo,
        string customer,
        string origin,
        string destination,
        decimal loaded,
        decimal received,
        decimal rate,
        decimal roadMoney,
        decimal otherExpense,
        OutputLayout layout,
        decimal bon = 0,
        decimal rejection = 0,
        decimal claim = 0) => new(
            date, plate, cargo, customer, origin, destination,
            loaded, received, rate, bon, rejection, claim,
            roadMoney, otherExpense,
            $"{DemoTag} Data percobaan sintetis",
            layout);
}
