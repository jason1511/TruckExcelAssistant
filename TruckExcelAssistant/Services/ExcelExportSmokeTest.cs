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
            var agrico = Path.Combine(directory, "invoice-agrico.xlsx");
            var complete = Path.Combine(directory, "invoice-lengkap.xlsx");
            var ledger = Path.Combine(directory, "pembukuan.xlsx");
            var settings = new AppSettings(
                "PT TEST TRANSPORT", "Jl. Contoh 1", "Lumajang", "BCA", "1234567890",
                "PT TEST TRANSPORT", "TEST SIGNER", "TJ", 3,
                OutputLayout.SumberPanganLike, directory);

            exporter.ExportCompactInvoice(records, "PT CONTOH CUSTOMER", "TJ-20260903-001", DateTime.Today, compact, settings);
            exporter.ExportAgricoInvoice(
                [records[0] with { Draft = records[0].Draft with { ClaimWeightKg = 107, ClaimRatePerKg = 6_950, Layout = OutputLayout.AgricoLike } }],
                "PT AGRICO TEST", "TJ-20260903-002", DateTime.Today, agrico, settings);
            exporter.ExportCompleteInvoice(records, "PT CONTOH CUSTOMER", "TJ-20260903-002", DateTime.Today, complete, settings);
            exporter.ExportTruckLedger(records, ledger, expenses);

            Verify(compact, "INV", "J4");
            VerifyAgrico(agrico);
            Verify(complete, "Invoice", "K19", "A21", "PT TEST TRANSPORT");
            Verify(ledger, "N-TEST-01", "A31", "H5", "Ban: Ganti ban belakang");
            VerifyFromToColumns(complete, ledger);
            VerifyInvoiceDatabase(directory, compact);
        }
        finally
        {
            Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
            Directory.Delete(directory, true);
        }
    }

    private static void VerifyFromToColumns(string completePath, string ledgerPath)
    {
        using (var workbook = new XLWorkbook(completePath))
        {
            var sheet = workbook.Worksheet("Invoice");
            if (sheet.Cell("I2").GetString() != "DARI"
                || sheet.Cell("J2").GetString() != "TUJUAN"
                || sheet.Cell("I4").GetString() != "Jember"
                || sheet.Cell("J4").GetString() != "Cirebon"
                || sheet.Cell("K4").GetDouble() != 12_521_250
                || sheet.Cell("K18").GetDouble() != 25_000)
            {
                throw new InvalidOperationException("Mirip Sumber Pangan tidak menggunakan kolom Dari dan Ke dengan benar.");
            }
        }

        using (var workbook = new XLWorkbook(ledgerPath))
        {
            var sheet = workbook.Worksheet("N-TEST-01");
            if (sheet.Cell("B4").GetString() != "DARI"
                || sheet.Cell("C4").GetString() != "KE"
                || sheet.Cell("B6").GetString() != "Jember"
                || sheet.Cell("C6").GetString() != "Cirebon")
            {
                throw new InvalidOperationException("Pembukuan tidak menggunakan kolom Dari dan Ke dengan benar.");
            }
        }
    }

    private static void VerifyAgrico(string path)
    {
        using var workbook = new XLWorkbook(path);
        var invoice = workbook.Worksheets.Single(sheet => sheet.Name.StartsWith("INV", StringComparison.Ordinal));
        var claim = workbook.Worksheets.Single(sheet => sheet.Name.StartsWith("KLAIM", StringComparison.Ordinal));
        if (workbook.Worksheets.Count != 2
            || invoice.Cell("L4").FormulaA1 != "H4+I4"
            || claim.Cell("I5").FormulaA1 != "G5*H5"
            || claim.Cell("G5").GetDouble() != 107
            || claim.Cell("H5").GetDouble() != 6_950)
        {
            throw new InvalidOperationException("Mirip Agrico tidak membuat pasangan sheet INV dan KLAIM dengan benar.");
        }
    }

    private static void VerifyInvoiceDatabase(string directory, string invoicePath)
    {
        var database = new DatabaseService(Path.Combine(directory, "smoke-test.db"));
        database.Initialize();
        if (OutputLayoutNames.ParseInvoiceSetting("CompactInvoice", OutputLayout.SumberPanganLike) != OutputLayout.MigunoLike
            || OutputLayoutNames.ParseInvoiceSetting("CompleteInvoice", OutputLayout.MigunoLike) != OutputLayout.SumberPanganLike)
        {
            throw new InvalidOperationException("Nama layout lama tidak dimigrasikan dengan benar.");
        }
        var date = new DateTime(2026, 9, 3);
        if (database.GetNextInvoiceNumber(date) != "TJ-20260903-001")
        {
            throw new InvalidOperationException("Nomor invoice otomatis pertama tidak sesuai.");
        }
        database.RecordGeneratedInvoice(
            "TJ-20260903-001",
            date,
            "PT CONTOH CUSTOMER",
            OutputLayout.SumberPanganLike,
            12_500_000,
            invoicePath,
            [],
            70_000);
        if (database.GetNextInvoiceNumber(date) != "TJ-20260903-002")
        {
            throw new InvalidOperationException("Urutan nomor invoice otomatis tidak bertambah.");
        }
        var invoices = database.GetInvoices();
        if (invoices.Count != 1
            || invoices[0].Status != InvoiceStatus.Generated
            || invoices[0].ClaimAmount != 70_000)
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
            OutputLayout.MigunoLike, directory));
        var settings = database.GetSettings();
        if (settings.CompanyName != "PT TEST TRANSPORT"
            || settings.DefaultInvoiceLayout != OutputLayout.MigunoLike
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
        var agricoDraft = SampleRecords()[0].Draft with
        {
            LicencePlate = "N-AGRICO-01",
            Customer = "PT AGRICO TEST",
            ClaimWeightKg = 107,
            ClaimRatePerKg = 6_950,
            ClaimAmount = 0,
            RejectionCost = 0,
            Layout = OutputLayout.AgricoLike
        };
        var agricoId = database.AddHaul(agricoDraft, HaulStatus.Saved);
        var agrico = database.GetHaul(agricoId)?.Draft;
        if (agrico is null
            || agrico.Layout != OutputLayout.AgricoLike
            || agrico.ClaimWeightKg != 107
            || agrico.ClaimRatePerKg != 6_950
            || agrico.EffectiveClaimAmount != 743_650)
        {
            throw new InvalidOperationException("Data entry Mirip Agrico tidak tersimpan atau dihitung dengan benar.");
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
            sheet.Cell("H2").Value = "JUMLAH";
            sheet.Cell("I2").Value = "DARI";
            sheet.Cell("J2").Value = "TUJUAN";
            sheet.Cell("K2").Value = "TOTAL";
            sheet.Cell("A4").Value = 1;
            sheet.Cell("B4").Value = date;
            sheet.Cell("C4").Value = "Jagung";
            sheet.Cell("D4").Value = "N-LEGACY-01";
            sheet.Cell("E4").Value = 46_000;
            sheet.Cell("F4").Value = 45_000;
            sheet.Cell("G4").Value = 300;
            sheet.Cell("H4").Value = 13_500_000;
            sheet.Cell("I4").Value = "Lumajang";
            sheet.Cell("J4").Value = "Semarang";
            sheet.Cell("K4").Value = 13_650_000;
            sheet.Cell("A18").Value = "KLAIM";
            sheet.Cell("K18").Value = 70_000;
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
            || imported.Draft.Destination != "Semarang"
            || imported.Draft.RejectionCost != 150_000
            || imported.Draft.ClaimAmount != 70_000
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

        var demo = new DemoDataSeeder(database, new ExcelExportService());
        demo.Seed();
        if (!database.HasDemoData()
            || database.CountHauls() != 15
            || database.GetExpenses().Count != 8
            || database.GetInvoices().Count != 3
            || database.GetInvoices(status: InvoiceStatus.Generated).Count != 1)
        {
            throw new InvalidOperationException("Data contoh tidak dibuat dengan lengkap.");
        }
        demo.Remove();
        if (database.HasDemoData()
            || database.CountHauls() != 3
            || database.GetExpenses().Count != 2
            || database.GetInvoices().Count != 1)
        {
            throw new InvalidOperationException("Data contoh tidak dapat dibersihkan dengan aman.");
        }
    }

    private static IReadOnlyList<HaulRecord> SampleRecords()
    {
        var now = DateTime.UtcNow;
        return
        [
            new HaulRecord(1, new HaulDraft(
                new DateTime(2026, 9, 1), "N-TEST-01", "Jagung", "PT CONTOH CUSTOMER",
                "Jember", "Cirebon", 45_500, 45_350, 275, 100_000, 50_000, 0, 0, 25_000,
                4_000_000, 250_000, "Biaya tol", OutputLayout.SumberPanganLike),
                HaulStatus.Saved, now, now, null),
            new HaulRecord(2, new HaulDraft(
                new DateTime(2026, 9, 2), "N-TEST-01", "SBM", "PT CONTOH CUSTOMER",
                "Surabaya", "Semarang", 44_000, 43_900, 145, 0, 0, 0, 0, 0,
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
        var sheet = workbook.Worksheets.FirstOrDefault(item => item.Name.StartsWith(expectedSheet, StringComparison.Ordinal));
        if (workbook.Worksheets.Count != 1 || sheet is null)
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
