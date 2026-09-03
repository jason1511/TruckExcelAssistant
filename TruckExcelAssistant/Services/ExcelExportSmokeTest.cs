using ClosedXML.Excel;
using TruckExcelAssistant.Models;

namespace TruckExcelAssistant.Services;

internal static class ExcelExportSmokeTest
{
    public static void Run()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"truck-excel-smoke-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            var records = SampleRecords();
            var expenses = SampleExpenses();
            var exporter = new ExcelExportService();
            var compact = Path.Combine(directory, "invoice-ringkas.xlsx");
            var complete = Path.Combine(directory, "invoice-lengkap.xlsx");
            var ledger = Path.Combine(directory, "pembukuan.xlsx");
            var settings = new AppSettings(
                "PT TEST TRANSPORT", "Jl. Contoh 1", "Lumajang", "BCA", "1234567890",
                "PT TEST TRANSPORT", "TEST SIGNER", "TJ", 3,
                OutputLayout.CompleteInvoice, directory);

            exporter.ExportCompactInvoice(records, "PT CONTOH CUSTOMER", "TJ-20260903-001", DateTime.Today, compact, settings);
            exporter.ExportCompleteInvoice(records, "PT CONTOH CUSTOMER", "TJ-20260903-002", DateTime.Today, complete, settings);
            exporter.ExportTruckLedger(records, ledger, expenses);

            Verify(compact, "Invoice", "I4", "A24", "PT TEST TRANSPORT");
            Verify(complete, "Invoice", "K19", "A21", "PT TEST TRANSPORT");
            Verify(ledger, "N-TEST-01", "A121", "H5", "Ban: Ganti ban belakang");
            VerifyInvoiceDatabase(directory, compact);
        }
        finally
        {
            Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
            Directory.Delete(directory, true);
        }
    }

    private static void VerifyInvoiceDatabase(string directory, string invoicePath)
    {
        var database = new DatabaseService(Path.Combine(directory, "smoke-test.db"));
        database.Initialize();
        var date = new DateTime(2026, 9, 3);
        if (database.GetNextInvoiceNumber(date) != "TJ-20260903-001")
        {
            throw new InvalidOperationException("Nomor invoice otomatis pertama tidak sesuai.");
        }
        database.RecordGeneratedInvoice(
            "TJ-20260903-001",
            date,
            "PT CONTOH CUSTOMER",
            OutputLayout.CompleteInvoice,
            12_500_000,
            invoicePath,
            []);
        if (database.GetNextInvoiceNumber(date) != "TJ-20260903-002")
        {
            throw new InvalidOperationException("Urutan nomor invoice otomatis tidak bertambah.");
        }
        var invoices = database.GetInvoices();
        if (invoices.Count != 1 || invoices[0].Status != InvoiceStatus.Generated)
        {
            throw new InvalidOperationException("Riwayat invoice tidak tersimpan dengan benar.");
        }
        database.UpdateInvoiceStatus(invoices[0].Id, InvoiceStatus.Paid);
        if (database.GetInvoices(status: InvoiceStatus.Paid).Count != 1)
        {
            throw new InvalidOperationException("Status lunas invoice tidak tersimpan.");
        }
        database.SaveSettings(new AppSettings(
            "PT TEST TRANSPORT", "Jl. Contoh 1", "Jember", "BCA", "1234567890",
            "PT TEST TRANSPORT", "TEST SIGNER", "TJ", 4,
            OutputLayout.CompactInvoice, directory));
        var settings = database.GetSettings();
        if (settings.CompanyName != "PT TEST TRANSPORT"
            || settings.DefaultInvoiceLayout != OutputLayout.CompactInvoice
            || database.GetNextInvoiceNumber(date) != "TJ-20260903-0002")
        {
            throw new InvalidOperationException("Pengaturan invoice tidak tersimpan atau diterapkan.");
        }
        var expenseId = database.AddExpense(
            date, "N-TEST-01", "Ban", "Ganti ban belakang", 500_000);
        var expenses = database.GetExpenses();
        if (expenses.Count != 1 || expenses[0].Amount != 500_000)
        {
            throw new InvalidOperationException("Pengeluaran tidak tersimpan.");
        }
        database.UpdateExpense(expenseId, date, "N-TEST-01", "Ban", "Ganti dua ban", 750_000);
        if (database.GetExpensesForExport(date, date, "N-TEST-01").Single().Amount != 750_000)
        {
            throw new InvalidOperationException("Perubahan pengeluaran tidak tersimpan.");
        }
        database.MoveExpenseToTrash(expenseId);
        if (database.GetExpenses(deletedOnly: true).Count != 1)
        {
            throw new InvalidOperationException("Pengeluaran tidak masuk ke Sampah.");
        }
        database.RestoreExpenseFromTrash(expenseId);
        if (database.GetExpenses().Count != 1)
        {
            throw new InvalidOperationException("Pengeluaran tidak berhasil dipulihkan.");
        }
        database.AddHaul(SampleRecords()[0].Draft, HaulStatus.Saved);
        var dashboard = database.GetDashboardSummary(date);
        if (dashboard.HaulCount != 1
            || dashboard.Revenue != 12_471_250
            || dashboard.TotalExpenses != 5_000_000
            || dashboard.Net != 7_471_250
            || dashboard.OutstandingInvoiceCount != 0
            || dashboard.Trucks.Count != 1
            || dashboard.RecentInvoices.Count != 1)
        {
            throw new InvalidOperationException("Ringkasan bulanan tidak menghitung data dengan benar.");
        }
        VerifyLegacyImport(database, directory);
    }

    private static void VerifyLegacyImport(DatabaseService database, string directory)
    {
        var date = new DateTime(2025, 8, 1);
        var ledgerPath = Path.Combine(directory, "PEMBUKUAN TEST.xlsx");
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("N-LEGACY-01");
            sheet.Cell("A1").Value = "N-LEGACY-01";
            sheet.Cell("A2").Value = "TGL";
            sheet.Cell("B2").Value = "PEMASUKAN";
            sheet.Cell("H2").Value = "PENGELUARAN";
            sheet.Cell("B3").Value = "DARI";
            sheet.Cell("C3").Value = "KE";
            sheet.Cell("D3").Value = "BARANG";
            sheet.Cell("E3").Value = "BERAT (KG)";
            sheet.Cell("F3").Value = "ONGKOS";
            sheet.Cell("H3").Value = "KETERANGAN";
            sheet.Cell("I3").Value = "UANG JALAN SOPIR";
            sheet.Cell("J3").Value = "BIAYA";
            sheet.Cell("A5").Value = date;
            sheet.Cell("B5").Value = "Lumajang";
            sheet.Cell("C5").Value = "Semarang";
            sheet.Cell("D5").Value = "Jagung";
            sheet.Cell("E5").Value = 45_000;
            sheet.Cell("F5").Value = 300;
            sheet.Cell("I5").Value = 2_000_000;
            sheet.Cell("A6").Value = date;
            sheet.Cell("H6").Value = "Service radiator";
            sheet.Cell("J6").Value = 500_000;
            workbook.SaveAs(ledgerPath);
        }

        var invoicePath = Path.Combine(directory, "INVOICE TEST.xlsx");
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("INV 001");
            sheet.Cell("A1").Value = "Kepada Yth. PT CUSTOMER LEGACY";
            sheet.Cell("A2").Value = "NO.";
            sheet.Cell("B2").Value = "TANGGAL";
            sheet.Cell("C2").Value = "JENIS MUATAN";
            sheet.Cell("D2").Value = "NOPOL";
            sheet.Cell("E2").Value = "BERAT";
            sheet.Cell("F2").Value = "BERAT";
            sheet.Cell("G2").Value = "ONGKOS";
            sheet.Cell("A4").Value = 1;
            sheet.Cell("B4").Value = date;
            sheet.Cell("C4").Value = "Jagung";
            sheet.Cell("D4").Value = "N-LEGACY-01";
            sheet.Cell("E4").Value = 46_000;
            sheet.Cell("F4").Value = 45_000;
            sheet.Cell("G4").Value = 300;
            workbook.SaveAs(invoicePath);
        }

        var importer = new LegacyWorkbookImporter(database);
        var result = importer.Import([invoicePath, ledgerPath]);
        var imported = database.GetHauls("N-LEGACY-01").Single();
        if (result.AddedHauls != 1
            || result.UpdatedHauls != 1
            || result.AddedExpenses != 1
            || imported.Draft.Customer != "PT CUSTOMER LEGACY"
            || imported.Draft.Origin != "Lumajang"
            || imported.Draft.DriverRoadMoney != 2_000_000
            || imported.Draft.LoadedWeightKg != 46_000)
        {
            throw new InvalidOperationException("Impor dan penggabungan Excel lama tidak sesuai.");
        }
        var repeated = new LegacyWorkbookImporter(database).Import([ledgerPath, invoicePath]);
        if (repeated.SkippedRows != 3 || database.GetHauls("N-LEGACY-01").Count != 1)
        {
            throw new InvalidOperationException("Impor Excel lama membuat baris ganda.");
        }
    }

    private static IReadOnlyList<HaulRecord> SampleRecords()
    {
        var now = DateTime.UtcNow;
        return
        [
            new HaulRecord(1, new HaulDraft(
                new DateTime(2026, 9, 1), "N-TEST-01", "Jagung", "PT CONTOH CUSTOMER",
                "Jember", "Cirebon", 45_500, 45_350, 275, 100_000, 50_000, 25_000,
                4_000_000, 250_000, "Biaya tol", OutputLayout.CompleteInvoice),
                HaulStatus.Saved, now, now, null),
            new HaulRecord(2, new HaulDraft(
                new DateTime(2026, 9, 2), "N-TEST-01", "SBM", "PT CONTOH CUSTOMER",
                "Surabaya", "Semarang", 44_000, 43_900, 145, 0, 0, 0,
                3_000_000, 100_000, "", OutputLayout.TruckLedger),
                HaulStatus.Saved, now, now, null)
        ];
    }

    private static IReadOnlyList<ExpenseRecord> SampleExpenses()
    {
        var now = DateTime.UtcNow;
        return
        [
            new ExpenseRecord(
                1, new DateTime(2026, 8, 31), "N-TEST-01", "Ban",
                "Ganti ban belakang", 500_000, now, now, null)
        ];
    }

    private static void Verify(
        string path,
        string expectedSheet,
        string formulaCell,
        string? settingsCell = null,
        string? expectedSettingsValue = null)
    {
        if (!File.Exists(path) || new FileInfo(path).Length == 0)
        {
            throw new InvalidOperationException($"Smoke test tidak menghasilkan {Path.GetFileName(path)}.");
        }

        using var workbook = new XLWorkbook(path);
        if (workbook.Worksheets.Count != 1 || !workbook.TryGetWorksheet(expectedSheet, out var sheet))
        {
            throw new InvalidOperationException($"Sheet {expectedSheet} tidak ditemukan di {Path.GetFileName(path)}.");
        }
        if (string.IsNullOrWhiteSpace(sheet.Cell(formulaCell).FormulaA1))
        {
            throw new InvalidOperationException($"Formula {formulaCell} hilang dari {Path.GetFileName(path)}.");
        }
        if (settingsCell is not null && sheet.Cell(settingsCell).GetString() != expectedSettingsValue)
        {
            throw new InvalidOperationException($"Pengaturan tidak diterapkan ke {Path.GetFileName(path)}.");
        }
    }
}
