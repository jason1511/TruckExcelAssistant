using ClosedXML.Excel;
using TruckExcelAssistant.Models;

namespace TruckExcelAssistant.Services;

public sealed class ExcelExportService
{
    private static readonly int[] LedgerStartRows = [5, 35, 65, 95];
    private static readonly int[] LedgerTotalRows = [30, 60, 90, 120];

    public string ExportDirectory { get; } = Path.Combine(AppContext.BaseDirectory, "Exports");

    public void ExportCompactInvoice(
        IReadOnlyList<HaulRecord> records,
        string customer,
        string invoiceNumber,
        DateTime issueDate,
        string outputPath)
    {
        ValidateCount(records, 19, "Invoice ringkas");
        using var workbook = CreateCompactWorkbook();
        var sheet = workbook.Worksheet("Invoice");

        sheet.Cell("A1").Value = $" Kepada Yth.    :      {customer.Trim()}";
        sheet.Cell("I1").Value = invoiceNumber.Trim();
        for (var index = 0; index < records.Count; index++)
        {
            var row = index + 4;
            var haul = records[index].Draft;
            sheet.Cell(row, 1).Value = index + 1;
            sheet.Cell(row, 2).Value = haul.Date;
            sheet.Cell(row, 3).Value = haul.Cargo;
            sheet.Cell(row, 4).Value = haul.LicencePlate;
            SetNumber(sheet.Cell(row, 5), haul.ReceivedWeightKg);
            SetNumber(sheet.Cell(row, 6), haul.RatePerKg);
            sheet.Cell(row, 7).FormulaA1 = $"=E{row}*F{row}";
            SetNumber(sheet.Cell(row, 8), haul.BonSangu);
            sheet.Cell(row, 9).FormulaA1 = $"=G{row}-H{row}";
        }

        sheet.Cell("G23").FormulaA1 = "=SUM(G4:G22)";
        sheet.Cell("H23").FormulaA1 = "=SUM(H4:H22)";
        sheet.Cell("I23").FormulaA1 = "=SUM(I4:I22)";
        sheet.Cell("G24").Value = $"Lumajang, {IndonesianDate(issueDate)}";
        Finish(workbook, outputPath);
    }

    public void ExportCompleteInvoice(
        IReadOnlyList<HaulRecord> records,
        string customer,
        string invoiceNumber,
        DateTime issueDate,
        string outputPath)
    {
        ValidateCount(records, 13, "Invoice lengkap");
        using var workbook = CreateCompleteWorkbook();
        var sheet = workbook.Worksheet("Invoice");

        sheet.Cell("A1").Value = $"Kepada Yth. {customer.Trim()}";
        sheet.Cell("K1").Value = invoiceNumber.Trim();
        for (var index = 0; index < records.Count; index++)
        {
            var row = index + 4;
            var haul = records[index].Draft;
            sheet.Cell(row, 1).Value = index + 1;
            sheet.Cell(row, 2).Value = haul.Date;
            sheet.Cell(row, 3).Value = haul.Cargo;
            sheet.Cell(row, 4).Value = haul.LicencePlate;
            SetNumber(sheet.Cell(row, 5), haul.LoadedWeightKg);
            SetNumber(sheet.Cell(row, 6), haul.ReceivedWeightKg);
            SetNumber(sheet.Cell(row, 7), haul.RatePerKg);
            sheet.Cell(row, 8).FormulaA1 = $"=F{row}*G{row}";
            sheet.Cell(row, 9).Value = haul.Origin;
            sheet.Cell(row, 10).Value = haul.Destination;
            SetNumber(sheet.Cell(row, 11), haul.GrossAmount + haul.RejectionCost);
        }

        sheet.Cell("K17").FormulaA1 = "=SUM(K4:K16)";
        SetNumber(sheet.Cell("K18"), records.Sum(item => item.Draft.ClaimAmount));
        sheet.Cell("K19").FormulaA1 = "=K17-K18";
        sheet.Cell("I21").Value = $"Lumajang, {IndonesianDate(issueDate)}";
        Finish(workbook, outputPath);
    }

    public void ExportTruckLedger(IReadOnlyList<HaulRecord> records, string outputPath)
    {
        var groups = records
            .Where(record => !string.IsNullOrWhiteSpace(record.Draft.LicencePlate))
            .GroupBy(record => record.Draft.LicencePlate.Trim(), StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (groups.Count == 0)
        {
            throw new InvalidOperationException("Pilih setidaknya satu perjalanan yang memiliki nomor polisi.");
        }

        using var workbook = CreateLedgerWorkbook();
        var source = workbook.Worksheet("Truk");

        var sheets = new List<IXLWorksheet> { source };
        for (var index = 1; index < groups.Count; index++)
        {
            sheets.Add(source.CopyTo(SafeSheetName(groups[index].Key, workbook)));
        }

        for (var index = 0; index < groups.Count; index++)
        {
            var rows = groups[index].OrderBy(item => item.Draft.Date).ThenBy(item => item.Id).ToList();
            ValidateCount(rows, 100, $"Pembukuan {groups[index].Key}");
            var sheet = sheets[index];
            if (index == 0)
            {
                sheet.Name = SafeSheetName(groups[index].Key, workbook, sheet);
            }
            FillLedgerSheet(sheet, groups[index].Key, rows);
        }

        Finish(workbook, outputPath);
    }

    private static XLWorkbook CreateCompactWorkbook()
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Invoice");
        double[] widths = [7, 15, 18, 16, 15, 14, 19, 17, 19];
        SetColumnWidths(sheet, widths);
        sheet.Range("A1:H1").Merge();
        sheet.Cell("A1").Value = "Kepada Yth. :";
        sheet.Cell("I1").Value = "TJ";
        sheet.Row(1).Height = 24;

        string[] headers = ["NO.", "TGL\nMUAT", "JENIS\nMUATAN", "NOPOL", "BERAT\n(Kg.)", "ONGK\n(Rp./Kg)", "JUMLAH\n(Rp.)", "BON SANGU\n(Rp.)", "SISA ONGK.\n(Rp.)"];
        for (var column = 1; column <= headers.Length; column++)
        {
            sheet.Cell(2, column).Value = headers[column - 1];
        }
        sheet.Range("A2:I3").Style.Fill.BackgroundColor = XLColor.FromHtml("#FFD966");
        sheet.Range("A2:I3").Style.Font.Bold = true;
        sheet.Range("A2:I3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        sheet.Range("A2:I3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        sheet.Range("A2:I3").Style.Alignment.WrapText = true;
        sheet.Range("A2:A3").Merge();
        sheet.Range("B2:B3").Merge();
        sheet.Range("C2:C3").Merge();
        sheet.Range("D2:D3").Merge();
        sheet.Range("E2:E3").Merge();
        sheet.Range("F2:F3").Merge();
        sheet.Range("G2:G3").Merge();
        sheet.Range("H2:H3").Merge();
        sheet.Range("I2:I3").Merge();
        sheet.Row(2).Height = 36;
        sheet.Row(3).Height = 5;

        StyleDataArea(sheet.Range("A4:I22"), "#FFF2CC");
        sheet.Range("A4:I22").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        sheet.Range("E4:I23").Style.NumberFormat.Format = "#,##0";
        sheet.Range("B4:B22").Style.DateFormat.Format = "dd/MM/yyyy";
        sheet.Range("A23:F23").Merge();
        sheet.Cell("A23").Value = "TOTAL";
        sheet.Range("A23:I23").Style.Fill.BackgroundColor = XLColor.FromHtml("#806000");
        sheet.Range("A23:I23").Style.Font.FontColor = XLColor.White;
        sheet.Range("A23:I23").Style.Font.Bold = true;
        sheet.Range("A23:I23").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        sheet.Range("G24:I24").Merge();
        sheet.Cell("B26").Value = "Informasi pembayaran";
        sheet.Cell("B26").Style.Font.Italic = true;
        sheet.Cell("G26").Value = "Dibuat oleh";
        sheet.Cell("G28").Value = "________________________";
        sheet.Range("G24:I28").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ApplyBorders(sheet.Range("A2:I23"));
        sheet.SheetView.FreezeRows(3);
        sheet.PageSetup.PrintAreas.Add("A1:I28");
        sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        return workbook;
    }

    private static XLWorkbook CreateCompleteWorkbook()
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Invoice");
        double[] widths = [6, 13, 17, 15, 13, 14, 12, 17, 15, 15, 18];
        SetColumnWidths(sheet, widths);
        sheet.Range("A1:J1").Merge();
        sheet.Cell("A1").Value = "Kepada Yth.";
        sheet.Cell("K1").Value = "TJ";
        sheet.Row(1).Height = 24;

        string[] headers = ["NO.", "TANGGAL", "JENIS\nMUATAN", "NOPOL", "BERAT\nMUAT", "BERAT\nDITERIMA", "ONGKOS", "JUMLAH", "DARI", "TUJUAN", "TOTAL"];
        for (var column = 1; column <= headers.Length; column++)
        {
            sheet.Cell(2, column).Value = headers[column - 1];
        }
        sheet.Range("A2:K3").Style.Fill.BackgroundColor = XLColor.FromHtml("#A9D18E");
        sheet.Range("A2:K3").Style.Font.Bold = true;
        sheet.Range("A2:K3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        sheet.Range("A2:K3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        sheet.Range("A2:K3").Style.Alignment.WrapText = true;
        for (var column = 1; column <= 11; column++)
        {
            sheet.Range(2, column, 3, column).Merge();
        }
        sheet.Row(2).Height = 36;
        sheet.Row(3).Height = 5;
        StyleDataArea(sheet.Range("A4:K16"), "#E2F0D9");
        sheet.Range("A4:K16").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        sheet.Range("E4:H19").Style.NumberFormat.Format = "#,##0";
        sheet.Range("K4:K19").Style.NumberFormat.Format = "#,##0";
        sheet.Range("B4:B16").Style.DateFormat.Format = "dd/MM/yyyy";
        foreach (var row in new[] { 17, 18, 19 })
        {
            sheet.Range(row, 1, row, 10).Merge();
            sheet.Range(row, 1, row, 11).Style.Fill.BackgroundColor = XLColor.FromHtml("#375623");
            sheet.Range(row, 1, row, 11).Style.Font.FontColor = XLColor.White;
            sheet.Range(row, 1, row, 11).Style.Font.Bold = true;
        }
        sheet.Cell("A17").Value = "JUMLAH";
        sheet.Cell("A18").Value = "KLAIM";
        sheet.Cell("A19").Value = "TOTAL";
        sheet.Range("A17:J19").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        sheet.Range("I21:K21").Merge();
        sheet.Cell("I23").Value = "Dibuat oleh";
        sheet.Cell("I27").Value = "________________________";
        sheet.Range("I21:K27").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ApplyBorders(sheet.Range("A2:K19"));
        sheet.SheetView.FreezeRows(3);
        sheet.PageSetup.PrintAreas.Add("A1:K27");
        sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        return workbook;
    }

    private static XLWorkbook CreateLedgerWorkbook()
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Truk");
        double[] widths = [13, 18, 18, 17, 15, 13, 18, 30, 17, 17];
        SetColumnWidths(sheet, widths);
        for (var block = 0; block < LedgerStartRows.Length; block++)
        {
            var titleRow = LedgerStartRows[block] - 4;
            var sectionRow = titleRow + 1;
            var headerRow = titleRow + 2;
            var subheaderRow = titleRow + 3;
            sheet.Cell(titleRow, 1).Value = "NOMOR POLISI";
            sheet.Range(titleRow, 1, titleRow, 10).Style.Font.Bold = true;
            sheet.Range(sectionRow, 2, sectionRow, 7).Merge();
            sheet.Cell(sectionRow, 2).Value = "PEMASUKAN";
            sheet.Range(sectionRow, 8, sectionRow, 10).Merge();
            sheet.Cell(sectionRow, 8).Value = "PENGELUARAN";
            sheet.Range(headerRow, 2, headerRow, 3).Merge();
            sheet.Cell(headerRow, 2).Value = "TUJUAN";
            sheet.Range(headerRow, 4, headerRow, 5).Merge();
            sheet.Cell(headerRow, 4).Value = "MUATAN";
            sheet.Cell(subheaderRow, 1).Value = "TGL";
            sheet.Cell(subheaderRow, 2).Value = "DARI";
            sheet.Cell(subheaderRow, 3).Value = "KE";
            sheet.Cell(subheaderRow, 4).Value = "BARANG";
            sheet.Cell(subheaderRow, 5).Value = "BERAT (KG)";
            sheet.Cell(headerRow, 6).Value = "ONGKOS";
            sheet.Cell(headerRow, 7).Value = "TOTAL";
            sheet.Cell(headerRow, 8).Value = "KETERANGAN";
            sheet.Cell(headerRow, 9).Value = "UANG JALAN";
            sheet.Cell(subheaderRow, 9).Value = "SOPIR";
            sheet.Cell(headerRow, 10).Value = "BIAYA";
            sheet.Range(sectionRow, 1, subheaderRow, 10).Style.Fill.BackgroundColor = XLColor.FromHtml("#D9EAD3");
            sheet.Range(sectionRow, 1, subheaderRow, 10).Style.Font.Bold = true;
            sheet.Range(sectionRow, 1, subheaderRow, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range(sectionRow, 1, subheaderRow, 10).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Range(sectionRow, 1, subheaderRow, 10).Style.Alignment.WrapText = true;
            StyleDataArea(sheet.Range(LedgerStartRows[block], 1, LedgerStartRows[block] + 24, 10), "#FFFFFF");
            var totalRow = LedgerTotalRows[block];
            sheet.Range(totalRow, 1, totalRow, 10).Style.Fill.BackgroundColor = XLColor.FromHtml("#E2F0D9");
            sheet.Range(totalRow, 1, totalRow, 10).Style.Font.Bold = true;
            ApplyBorders(sheet.Range(sectionRow, 1, totalRow, 10));
        }
        sheet.Range("A121:J121").Style.Fill.BackgroundColor = XLColor.FromHtml("#375623");
        sheet.Range("A121:J121").Style.Font.FontColor = XLColor.White;
        sheet.Range("A121:J121").Style.Font.Bold = true;
        sheet.Range("A5:A119").Style.DateFormat.Format = "dd/MM/yyyy";
        sheet.Range("E5:G121").Style.NumberFormat.Format = "#,##0";
        sheet.Range("I5:J121").Style.NumberFormat.Format = "#,##0";
        sheet.SheetView.FreezeRows(4);
        sheet.PageSetup.PrintAreas.Add("A1:J121");
        sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        return workbook;
    }

    private static void SetColumnWidths(IXLWorksheet sheet, IReadOnlyList<double> widths)
    {
        for (var index = 0; index < widths.Count; index++)
        {
            sheet.Column(index + 1).Width = widths[index];
        }
    }

    private static void StyleDataArea(IXLRange range, string alternatingColor)
    {
        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        for (var row = 1; row <= range.RowCount(); row++)
        {
            if (row % 2 == 1)
            {
                range.Row(row).Style.Fill.BackgroundColor = XLColor.FromHtml(alternatingColor);
            }
        }
    }

    private static void ApplyBorders(IXLRange range)
    {
        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        range.Style.Border.OutsideBorderColor = XLColor.FromHtml("#B7B7B7");
        range.Style.Border.InsideBorderColor = XLColor.FromHtml("#D9D9D9");
    }

    private static void FillLedgerSheet(IXLWorksheet sheet, string plate, IReadOnlyList<HaulRecord> records)
    {
        for (var block = 0; block < LedgerStartRows.Length; block++)
        {
            sheet.Cell(LedgerStartRows[block] - 4, 1).Value = plate;
            ClearContents(sheet, $"A{LedgerStartRows[block]}:J{LedgerStartRows[block] + 24}");
            var totalRow = LedgerTotalRows[block];
            sheet.Cell(totalRow, 7).FormulaA1 = $"=SUM(G{LedgerStartRows[block]}:G{totalRow - 1})";
            sheet.Cell(totalRow, 9).FormulaA1 = $"=SUM(I{LedgerStartRows[block]}:I{totalRow - 1})";
            sheet.Cell(totalRow, 10).FormulaA1 = $"=SUM(J{LedgerStartRows[block]}:J{totalRow - 1})";
        }

        for (var index = 0; index < records.Count; index++)
        {
            var block = index / 25;
            var row = LedgerStartRows[block] + index % 25;
            var haul = records[index].Draft;
            sheet.Cell(row, 1).Value = haul.Date;
            sheet.Cell(row, 2).Value = haul.Origin;
            sheet.Cell(row, 3).Value = haul.Destination;
            sheet.Cell(row, 4).Value = haul.Cargo;
            SetNumber(sheet.Cell(row, 5), haul.ReceivedWeightKg);
            SetNumber(sheet.Cell(row, 6), haul.RatePerKg);
            sheet.Cell(row, 7).FormulaA1 = $"=E{row}*F{row}";
            sheet.Cell(row, 8).Value = haul.Notes;
            SetNumber(sheet.Cell(row, 9), haul.DriverRoadMoney);
            SetNumber(sheet.Cell(row, 10), haul.OtherExpense);
        }

        sheet.Cell("G121").FormulaA1 = "=SUM(G30,G60,G90,G120)";
        sheet.Cell("I121").FormulaA1 = "=SUM(I30,I60,I90,I120)";
        sheet.Cell("J121").FormulaA1 = "=SUM(J30,J60,J90,J120)";
        sheet.Cell("A121").FormulaA1 = "=G121-I121-J121";
    }

    private static void ClearContents(IXLWorksheet sheet, string range) =>
        sheet.Range(range).Clear(XLClearOptions.Contents);

    private static void SetNumber(IXLCell cell, decimal value) => cell.Value = Convert.ToDouble(value);

    private static void ValidateCount<T>(IReadOnlyCollection<T> records, int maximum, string layout)
    {
        if (records.Count == 0)
        {
            throw new InvalidOperationException("Pilih setidaknya satu perjalanan untuk diekspor.");
        }
        if (records.Count > maximum)
        {
            throw new InvalidOperationException($"{layout} hanya memuat {maximum} baris. Kurangi pilihan data lalu coba lagi.");
        }
    }

    private static string IndonesianDate(DateTime date)
    {
        string[] months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        return $"{date.Day} {months[date.Month - 1]} {date.Year}";
    }

    private static string SafeSheetName(string value, XLWorkbook workbook, IXLWorksheet? current = null)
    {
        var invalid = new HashSet<char>(['[', ']', ':', '*', '?', '/', '\\']);
        var cleaned = new string(value.Where(character => !invalid.Contains(character)).ToArray()).Trim();
        cleaned = string.IsNullOrWhiteSpace(cleaned) ? "Truk" : cleaned;
        cleaned = cleaned.Length > 31 ? cleaned[..31] : cleaned;
        var candidate = cleaned;
        var suffix = 2;
        while (workbook.Worksheets.Any(sheet => sheet != current && sheet.Name.Equals(candidate, StringComparison.OrdinalIgnoreCase)))
        {
            var ending = $" {suffix++}";
            candidate = cleaned[..Math.Min(cleaned.Length, 31 - ending.Length)] + ending;
        }
        return candidate;
    }

    private static void Finish(XLWorkbook workbook, string outputPath)
    {
        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }
        workbook.SaveAs(outputPath);
    }
}
