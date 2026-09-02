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
            var exporter = new ExcelExportService();
            var compact = Path.Combine(directory, "invoice-ringkas.xlsx");
            var complete = Path.Combine(directory, "invoice-lengkap.xlsx");
            var ledger = Path.Combine(directory, "pembukuan.xlsx");

            exporter.ExportCompactInvoice(records, "PT CONTOH CUSTOMER", "TEST-001", DateTime.Today, compact);
            exporter.ExportCompleteInvoice(records, "PT CONTOH CUSTOMER", "TEST-002", DateTime.Today, complete);
            exporter.ExportTruckLedger(records, ledger);

            Verify(compact, "Invoice", "I4");
            Verify(complete, "Invoice", "K19");
            Verify(ledger, "N-TEST-01", "A121");
        }
        finally
        {
            Directory.Delete(directory, true);
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

    private static void Verify(string path, string expectedSheet, string formulaCell)
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
    }
}
