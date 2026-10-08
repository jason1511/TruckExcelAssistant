using ClosedXML.Excel;
using TruckExcelAssistant.Models;

namespace TruckExcelAssistant.Services;

public sealed class ExcelExportService
{
    private const int MigunoPageRows = 25;
    private const int AgricoPageRows = 22;
    private const int LedgerBlockRows = 25;
    public string ExportDirectory { get; } = Path.Combine(AppContext.BaseDirectory, "Exports");

    public void ExportCompactInvoice(IReadOnlyList<HaulRecord> records, string customer, string invoiceNumber,
        DateTime issueDate, string outputPath, AppSettings? settings = null)
    {
        RequireRecords(records);
        using var book = new XLWorkbook();
        BuildMiguno(book.AddWorksheet(SheetName("INV", invoiceNumber)), records, customer, invoiceNumber,
            issueDate, settings ?? AppSettings.Default);
        Save(book, outputPath);
    }

    public void ExportAgricoInvoice(IReadOnlyList<HaulRecord> records, string customer, string invoiceNumber,
        DateTime issueDate, string outputPath, AppSettings? settings = null)
    {
        RequireRecords(records);
        using var book = new XLWorkbook();
        var config = settings ?? AppSettings.Default;
        BuildAgricoInvoice(book.AddWorksheet(SheetName("INV", invoiceNumber)), records, customer, invoiceNumber, issueDate, config);
        BuildAgricoClaim(book.AddWorksheet(SheetName("KLAIM", invoiceNumber)), records);
        Save(book, outputPath);
    }

    public void ExportCompleteInvoice(IReadOnlyList<HaulRecord> records, string customer, string invoiceNumber,
        DateTime issueDate, string outputPath, AppSettings? settings = null, decimal? invoiceClaimAmount = null)
    {
        RequireMaximum(records, 13, "Mirip Sumber Pangan");
        using var book = new XLWorkbook();
        BuildSumberPangan(book.AddWorksheet("Invoice"), records, customer, invoiceNumber, issueDate,
            settings ?? AppSettings.Default, invoiceClaimAmount ?? 0);
        Save(book, outputPath);
    }

    public void ExportTruckLedger(IReadOnlyList<HaulRecord> records, string outputPath,
        IReadOnlyList<ExpenseRecord>? expenses = null)
    {
        expenses ??= [];
        var plates = records.Select(x => x.Draft.LicencePlate.Trim()).Concat(expenses.Select(x => x.LicencePlate.Trim()))
            .Where(x => x.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x).ToList();
        if (plates.Count == 0) throw new InvalidOperationException("Tidak ada perjalanan atau pengeluaran bernomor polisi untuk diekspor.");
        using var book = new XLWorkbook();
        foreach (var plate in plates)
        {
            var entries = records.Where(x => x.Draft.LicencePlate.Equals(plate, StringComparison.OrdinalIgnoreCase))
                .Select(x => new LedgerEntry(x.Draft.Date, 0, x.Id, x, null))
                .Concat(expenses.Where(x => x.LicencePlate.Equals(plate, StringComparison.OrdinalIgnoreCase))
                    .Select(x => new LedgerEntry(x.Date, 1, x.Id, null, x)))
                .OrderBy(x => x.Date).ThenBy(x => x.KindOrder).ThenBy(x => x.Id).ToList();
            BuildLedger(book.AddWorksheet(SafeSheetName(plate, book)), plate, entries,
                Math.Max(1, (int)Math.Ceiling(entries.Count / (double)LedgerBlockRows)));
        }
        Save(book, outputPath);
    }

    private static void BuildMiguno(IXLWorksheet sheet, IReadOnlyList<HaulRecord> records, string customer,
        string number, DateTime date, AppSettings settings)
    {
        Widths(sheet, [7, 15, 18, 16, 14, 15, 13, 18, 17, 19]);
        sheet.Range("A1:I1").Merge(); sheet.Cell("A1").Value = $"Kepada Yth.    :      {customer.Trim()}";
        sheet.Cell("A1").Style.Font.Bold = true; sheet.Cell("J1").Value = $"INV. NO. {number.Trim()}";
        sheet.Cell("J1").Style.Font.Bold = true; sheet.Cell("J1").Style.Font.FontColor = XLColor.Red;
        var rows = new List<int>(); var row = 2;
        for (var i = 0; i < records.Count; i++)
        {
            if (i % MigunoPageRows == 0) { MigunoHeader(sheet, row); row += 2; }
            var r = row++; rows.Add(r); var h = records[i].Draft;
            BaseValues(sheet, r, i + 1, h); sheet.Cell(r, 8).FormulaA1 = $"=F{r}*G{r}";
            Num(sheet.Cell(r, 9), h.BonSangu); sheet.Cell(r, 10).FormulaA1 = $"=H{r}-I{r}";
            DataStyle(sheet.Range(r, 1, r, 10), i, "#FFF2CC");
        }
        sheet.Range(row, 1, row, 4).Merge();
        foreach (var pair in new[] { (5, "E"), (6, "F"), (8, "H"), (10, "J") }) sheet.Cell(row, pair.Item1).FormulaA1 = Sum(rows, pair.Item2);
        TotalStyle(sheet.Range(row, 1, row, 10), "#806000");
        MergeWrite(sheet, row + 1, 2, 4, BankLocation(settings), true, true);
        MergeWrite(sheet, row + 2, 2, 4, settings.BankAccountHolder, true, true);
        MergeWrite(sheet, row + 3, 2, 4, settings.BankAccountNumber, true, true);
        MergeWrite(sheet, row + 2, 8, 10, $"{settings.City}, {IndonesianDate(date)}");
        if (settings.SignerName.Length > 0) MergeWrite(sheet, row + 4, 8, 10, settings.SignerName, true, true);
        FinishInvoice(sheet, row, 10, row + 4);
    }

    private static void MigunoHeader(IXLWorksheet sheet, int row)
    {
        HeaderValues(sheet, row,
            ["NO.", "TGL", "JENIS", "NOPOL", "BERAT", "BERAT", "ONGK", "JUMLAH", "BON SANGU", "SISA ONGK."],
            ["", "MUAT", "MUATAN", "", "MUAT", "DITERIMA", "( Rp./Kg )", "( Rp. )", "( Rp. )", "( Rp. )"]);
        sheet.Range(row, 1, row + 1, 1).Merge(); sheet.Range(row, 4, row + 1, 4).Merge();
        HeaderStyle(sheet.Range(row, 1, row + 1, 10), "#FFD966");
    }

    private static void BuildAgricoInvoice(IXLWorksheet sheet, IReadOnlyList<HaulRecord> records, string customer,
        string number, DateTime date, AppSettings settings)
    {
        Widths(sheet, [6, 13, 17, 15, 13, 14, 12, 17, 13, 15, 15, 18]);
        sheet.Range("A1:K1").Merge(); sheet.Cell("A1").Value = $"Kepada Yth. {customer.Trim()}";
        sheet.Cell("A1").Style.Font.Bold = true; sheet.Cell("L1").Value = $"INV  {number.Trim()}";
        sheet.Cell("L1").Style.Font.Bold = true;
        var rows = new List<int>(); var row = 2;
        for (var i = 0; i < records.Count; i++)
        {
            if (i % AgricoPageRows == 0) { if (i > 0) row += 4; AgricoHeader(sheet, row); row += 2; }
            var r = row++; rows.Add(r); var h = records[i].Draft;
            BaseValues(sheet, r, i + 1, h); sheet.Cell(r, 8).FormulaA1 = $"=F{r}*G{r}";
            Num(sheet.Cell(r, 9), h.RejectionCost); sheet.Cell(r, 10).Value = h.Origin;
            sheet.Cell(r, 11).Value = h.Destination; sheet.Cell(r, 12).FormulaA1 = $"=H{r}+I{r}";
            DataStyle(sheet.Range(r, 1, r, 12), i, "#E2F0D9");
        }
        var amount = row; var claim = row + 1; var total = row + 2;
        foreach (var r in new[] { amount, claim, total }) sheet.Range(r, 1, r, 11).Merge();
        sheet.Cell(amount, 1).Value = "JUMLAH"; sheet.Cell(claim, 1).Value = "KLAIM"; sheet.Cell(total, 1).Value = "TOTAL";
        sheet.Cell(amount, 12).FormulaA1 = Sum(rows, "L"); Num(sheet.Cell(claim, 12), records.Sum(x => x.Draft.EffectiveClaimAmount));
        sheet.Cell(total, 12).FormulaA1 = $"=L{amount}-L{claim}"; TotalStyle(sheet.Range(amount, 1, total, 12), "#375623");
        sheet.Range(amount, 1, total, 11).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        var footer = total + 2; MergeWrite(sheet, footer, 2, 7, PaymentRequest(settings));
        MergeWrite(sheet, footer + 1, 2, 7, AccountLine(settings));
        MergeWrite(sheet, footer, 10, 12, $"{settings.City}, {IndonesianDate(date)}");
        MergeWrite(sheet, footer + 2, 10, 12, "Dibuat oleh");
        MergeWrite(sheet, footer + 6, 10, 12, settings.SignerName.Length == 0 ? "________________________" : settings.SignerName, true, true);
        FinishInvoice(sheet, total, 12, footer + 6);
    }

    private static void AgricoHeader(IXLWorksheet sheet, int row)
    {
        HeaderValues(sheet, row,
            ["NO.", "TANGGAL", "JENIS", "NOPOL", "BERAT", "BERAT", "ONGKOS", "JUMLAH", "BIAYA", "DARI", "TUJUAN", "TOTAL"],
            ["", "", "MUATAN", "", "MUAT", "DITERIMA", "", "", "TOLAKAN", "", "", ""]);
        foreach (var c in new[] { 1, 2, 4, 7, 8, 10, 11, 12 }) sheet.Range(row, c, row + 1, c).Merge();
        HeaderStyle(sheet.Range(row, 1, row + 1, 12), "#A9D18E");
    }

    private static void BuildAgricoClaim(IXLWorksheet sheet, IReadOnlyList<HaulRecord> records)
    {
        Widths(sheet, [7, 14, 17, 14, 14, 12, 12, 14, 18]);
        HeaderValues(sheet, 3, ["NO.", "TANGGAL", "NOPOL", "BERAT", "BERAT", "SUSUT", "KLAIM", "HARGA", "JUMLAH"],
            ["", "", "", "MUAT", "DITERIMA", "", "", "", ""]);
        foreach (var c in new[] { 1, 2, 3, 6, 7, 8, 9 }) sheet.Range(3, c, 4, c).Merge();
        HeaderStyle(sheet.Range("A3:I4"), "#A9D18E");
        var claims = records.Where(x => x.Draft.EffectiveClaimAmount > 0 || x.Draft.ClaimWeightKg > 0 || x.Draft.ClaimRatePerKg > 0).ToList();
        var rows = new List<int>();
        for (var i = 0; i < claims.Count; i++)
        {
            var r = i + 5; rows.Add(r); var h = claims[i].Draft;
            sheet.Cell(r, 1).Value = i + 1; sheet.Cell(r, 2).Value = h.Date; sheet.Cell(r, 3).Value = h.LicencePlate;
            Num(sheet.Cell(r, 4), h.LoadedWeightKg); Num(sheet.Cell(r, 5), h.ReceivedWeightKg);
            sheet.Cell(r, 6).FormulaA1 = $"=D{r}-E{r}"; Num(sheet.Cell(r, 7), h.ClaimWeightKg); Num(sheet.Cell(r, 8), h.ClaimRatePerKg);
            if (h.CalculatedClaimAmount > 0) sheet.Cell(r, 9).FormulaA1 = $"=G{r}*H{r}"; else Num(sheet.Cell(r, 9), h.ClaimAmount);
            DataStyle(sheet.Range(r, 1, r, 9), i, "#E2F0D9");
        }
        var total = claims.Count + 5; sheet.Range(total, 1, total, 8).Merge(); sheet.Cell(total, 1).Value = "TOTAL KLAIM";
        sheet.Cell(total, 9).FormulaA1 = rows.Count == 0 ? "=0" : Sum(rows, "I"); TotalStyle(sheet.Range(total, 1, total, 9), "#375623");
        Borders(sheet.Range(3, 1, total, 9), XLColor.White); sheet.Range(5, 2, total, 2).Style.DateFormat.Format = "dd/MM/yyyy";
        sheet.Range(5, 4, total, 9).Style.NumberFormat.Format = "#,##0"; PrintSetup(sheet, $"A1:I{total}", 4);
    }

    private static void BuildSumberPangan(IXLWorksheet sheet, IReadOnlyList<HaulRecord> records, string customer,
        string number, DateTime date, AppSettings settings, decimal claim)
    {
        Widths(sheet, [6, 13, 17, 15, 13, 14, 12, 17, 15, 15, 18]);
        sheet.Range("A1:J1").Merge(); sheet.Cell("A1").Value = $"Kepada Yth. {customer.Trim()}"; sheet.Cell("K1").Value = $"INV  {number.Trim()}";
        string[] headers = ["NO.", "TANGGAL", "JENIS\nMUATAN", "NOPOL", "BERAT\nMUAT", "BERAT\nDITERIMA", "ONGKOS", "JUMLAH", "DARI", "TUJUAN", "TOTAL"];
        for (var c = 1; c <= 11; c++) { sheet.Cell(2, c).Value = headers[c - 1]; sheet.Range(2, c, 3, c).Merge(); }
        var header = sheet.Range("A2:K3"); header.Style.Font.Bold = true;
        header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        header.Style.Alignment.WrapText = true; Borders(header, XLColor.Black);
        for (var i = 0; i < records.Count; i++)
        {
            var r = i + 4; var h = records[i].Draft; BaseValues(sheet, r, i + 1, h);
            sheet.Cell(r, 8).FormulaA1 = $"=F{r}*G{r}"; sheet.Cell(r, 9).Value = h.Origin;
            sheet.Cell(r, 10).Value = h.Destination; Num(sheet.Cell(r, 11), h.GrossAmount + h.RejectionCost);
        }
        sheet.Range("A4:K16").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; Borders(sheet.Range("A4:K16"), XLColor.FromHtml("#D9D9D9"));
        foreach (var r in new[] { 17, 18, 19 }) sheet.Range(r, 1, r, 10).Merge();
        sheet.Cell("A17").Value = "JUMLAH"; sheet.Cell("A18").Value = "KLAIM"; sheet.Cell("A19").Value = "TOTAL";
        sheet.Cell("K17").FormulaA1 = "=SUM(K4:K16)"; Num(sheet.Cell("K18"), claim); sheet.Cell("K19").FormulaA1 = "=K17-K18";
        sheet.Range("A17:K19").Style.Font.Bold = true; sheet.Range("A17:J19").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        Borders(sheet.Range("A17:K19"), XLColor.Black);
        MergeWrite(sheet, 21, 1, 8, settings.CompanyName, true); MergeWrite(sheet, 22, 1, 8, settings.CompanyAddress);
        MergeWrite(sheet, 23, 1, 8, PaymentRequest(settings)); MergeWrite(sheet, 24, 1, 8, AccountLine(settings));
        MergeWrite(sheet, 21, 9, 11, $"{settings.City}, {IndonesianDate(date)}"); MergeWrite(sheet, 23, 9, 11, "Dibuat oleh");
        MergeWrite(sheet, 27, 9, 11, settings.SignerName.Length == 0 ? "________________________" : settings.SignerName);
        sheet.Range("B4:B16").Style.DateFormat.Format = "dd/MM/yyyy"; sheet.Range("E4:H19").Style.NumberFormat.Format = "#,##0";
        sheet.Range("K4:K19").Style.NumberFormat.Format = "#,##0"; PrintSetup(sheet, "A1:K27", 3, true);
    }

    private static void BuildLedger(IXLWorksheet sheet, string plate, IReadOnlyList<LedgerEntry> entries, int blocks)
    {
        Widths(sheet, [13, 18, 18, 17, 15, 13, 18, 30, 17, 17]); var subtotalRows = new List<int>();
        for (var b = 0; b < blocks; b++)
        {
            var title = b * 30 + 1; var section = title + 1; var header = title + 2; var sub = title + 3;
            var first = title + 4; var subtotal = title + 29; subtotalRows.Add(subtotal);
            sheet.Cell(title, 1).Value = plate; sheet.Cell(title, 1).Style.Font.Bold = true; sheet.Cell(title, 1).Style.Font.Italic = true;
            sheet.Range(section, 1, sub, 1).Merge(); sheet.Cell(section, 1).Value = "TGL";
            sheet.Range(section, 2, section, 7).Merge(); sheet.Cell(section, 2).Value = "PEMASUKAN";
            sheet.Range(section, 8, section, 10).Merge(); sheet.Cell(section, 8).Value = "PENGELUARAN";
            sheet.Range(header, 2, header, 3).Merge(); sheet.Cell(header, 2).Value = "TUJUAN";
            sheet.Range(header, 4, header, 5).Merge(); sheet.Cell(header, 4).Value = "MUATAN";
            Vertical(sheet, header, sub, 6, "ONGKOS"); Vertical(sheet, header, sub, 7, "TOTAL"); Vertical(sheet, header, sub, 8, "KETERANGAN");
            sheet.Cell(header, 9).Value = "UANG JALAN"; sheet.Cell(sub, 9).Value = "SOPIR"; Vertical(sheet, header, sub, 10, "BIAYA");
            sheet.Cell(sub, 2).Value = "DARI"; sheet.Cell(sub, 3).Value = "KE"; sheet.Cell(sub, 4).Value = "BARANG"; sheet.Cell(sub, 5).Value = "BERAT (KG)";
            LedgerHeader(sheet.Range(section, 1, sub, 7), "#B7CF7B"); LedgerHeader(sheet.Range(section, 8, sub, 10), "#AFC4E6");
            Borders(sheet.Range(first, 1, subtotal - 1, 10), XLColor.FromHtml("#D9D9D9"));
            sheet.Range(subtotal, 1, subtotal, 10).Style.Fill.BackgroundColor = XLColor.FromHtml("#7FD3C7");
            sheet.Cell(subtotal, 7).FormulaA1 = $"=SUM(G{first}:G{subtotal - 1})"; sheet.Cell(subtotal, 9).FormulaA1 = $"=SUM(I{first}:I{subtotal - 1})";
            sheet.Cell(subtotal, 10).FormulaA1 = $"=SUM(J{first}:J{subtotal - 1})"; Borders(sheet.Range(subtotal, 1, subtotal, 10), XLColor.FromHtml("#D9D9D9"));
        }
        for (var i = 0; i < entries.Count; i++)
        {
            var r = i / LedgerBlockRows * 30 + 5 + i % LedgerBlockRows; var e = entries[i]; sheet.Cell(r, 1).Value = e.Date;
            if (e.Haul is not null)
            {
                var h = e.Haul.Draft; sheet.Cell(r, 2).Value = h.Origin; sheet.Cell(r, 3).Value = h.Destination; sheet.Cell(r, 4).Value = h.Cargo;
                Num(sheet.Cell(r, 5), h.ReceivedWeightKg); Num(sheet.Cell(r, 6), h.RatePerKg); sheet.Cell(r, 7).FormulaA1 = $"=E{r}*F{r}";
                sheet.Cell(r, 8).Value = h.Notes; Num(sheet.Cell(r, 9), h.DriverRoadMoney); Num(sheet.Cell(r, 10), h.OtherExpense);
            }
            else if (e.Expense is not null)
            {
                var x = e.Expense; sheet.Cell(r, 8).Value = x.Description.Length == 0 ? x.Category : $"{x.Category}: {x.Description}";
                Num(sheet.Cell(r, x.Category.Equals("Uang jalan tambahan", StringComparison.OrdinalIgnoreCase) ? 9 : 10), x.Amount);
            }
        }
        var grand = blocks * 30 + 1;
        foreach (var pair in new[] { (7, "G"), (9, "I"), (10, "J") }) sheet.Cell(grand, pair.Item1).FormulaA1 = $"=SUM({string.Join(",", subtotalRows.Select(r => $"{pair.Item2}{r}"))})";
        sheet.Cell(grand, 1).FormulaA1 = $"=G{grand}-I{grand}-J{grand}"; sheet.Range(grand, 1, grand, 10).Style.Fill.BackgroundColor = XLColor.FromHtml("#7FD3C7");
        sheet.Range(grand, 1, grand, 10).Style.Font.Bold = true; Borders(sheet.Range(grand, 1, grand, 10), XLColor.FromHtml("#D9D9D9"));
        sheet.Range(1, 1, grand, 10).Style.Font.FontName = "Arial"; sheet.Range(1, 1, grand, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        sheet.Range(1, 8, grand, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; sheet.Range(1, 5, grand, 10).Style.NumberFormat.Format = "#,##0;[Red]-#,##0;-";
        sheet.Range(1, 1, grand, 1).Style.DateFormat.Format = "dd/MM/yy"; foreach (var b in Enumerable.Range(0, blocks)) sheet.Cell(b * 30 + 1, 1).Style.NumberFormat.Format = "@";
        PrintSetup(sheet, $"A1:J{grand}", 4);
    }

    private static void BaseValues(IXLWorksheet s, int r, int no, HaulDraft h)
    { s.Cell(r, 1).Value = no; s.Cell(r, 2).Value = h.Date; s.Cell(r, 3).Value = h.Cargo; s.Cell(r, 4).Value = h.LicencePlate; Num(s.Cell(r, 5), h.LoadedWeightKg); Num(s.Cell(r, 6), h.ReceivedWeightKg); Num(s.Cell(r, 7), h.RatePerKg); }
    private static void HeaderValues(IXLWorksheet s, int r, IReadOnlyList<string> a, IReadOnlyList<string> b)
    { for (var c = 1; c <= a.Count; c++) { s.Cell(r, c).Value = a[c - 1]; s.Cell(r + 1, c).Value = b[c - 1]; } }
    private static void HeaderStyle(IXLRange r, string color)
    { r.Style.Fill.BackgroundColor = XLColor.FromHtml(color); r.Style.Font.Bold = true; r.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; r.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center; r.Style.Alignment.WrapText = true; Borders(r, XLColor.White); }
    private static void LedgerHeader(IXLRange r, string color) => HeaderStyle(r, color);
    private static void DataStyle(IXLRange r, int i, string color)
    { if (i % 2 == 0) r.Style.Fill.BackgroundColor = XLColor.FromHtml(color); r.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; r.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center; }
    private static void TotalStyle(IXLRange r, string color)
    { r.Style.Fill.BackgroundColor = XLColor.FromHtml(color); r.Style.Font.Bold = true; r.Style.Font.FontColor = XLColor.White; r.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; r.Style.NumberFormat.Format = "#,##0"; }
    private static void Borders(IXLRange r, XLColor color)
    { r.Style.Border.OutsideBorder = XLBorderStyleValues.Thin; r.Style.Border.InsideBorder = XLBorderStyleValues.Thin; r.Style.Border.OutsideBorderColor = color; r.Style.Border.InsideBorderColor = color; }
    private static void MergeWrite(IXLWorksheet s, int r, int c1, int c2, string? value, bool bold = false, bool italic = false)
    { s.Range(r, c1, r, c2).Merge(); s.Cell(r, c1).Value = value ?? ""; s.Cell(r, c1).Style.Font.Bold = bold; s.Cell(r, c1).Style.Font.Italic = italic; s.Cell(r, c1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; }
    private static void Vertical(IXLWorksheet s, int r1, int r2, int c, string value) { s.Range(r1, c, r2, c).Merge(); s.Cell(r1, c).Value = value; }
    private static void FinishInvoice(IXLWorksheet s, int total, int lastColumn, int printLast)
    { Borders(s.Range(2, 1, total, lastColumn), XLColor.White); s.Range(4, 2, total, 2).Style.DateFormat.Format = "dd/MM/yyyy"; s.Range(4, 5, total, lastColumn).Style.NumberFormat.Format = "#,##0"; PrintSetup(s, $"A1:{XLHelper.GetColumnLetterFromNumber(lastColumn)}{printLast}", 3); }
    private static void PrintSetup(IXLWorksheet s, string area, int freeze, bool onePage = false)
    { s.PageSetup.PrintAreas.Add(area); s.PageSetup.PageOrientation = XLPageOrientation.Landscape; s.PageSetup.FitToPages(1, onePage ? 1 : 0); s.SheetView.FreezeRows(freeze); s.ShowGridLines = false; }
    private static void Widths(IXLWorksheet s, IReadOnlyList<double> widths) { for (var i = 0; i < widths.Count; i++) s.Column(i + 1).Width = widths[i]; }
    private static string Sum(IEnumerable<int> rows, string col) => $"=SUM({string.Join(",", rows.Select(r => $"{col}{r}"))})";
    private static void Num(IXLCell c, decimal value) => c.Value = Convert.ToDouble(value);
    private static string BankLocation(AppSettings s) => s.BankName.Length == 0 ? "" : $"{s.BankName} {s.City}".Trim();
    private static string PaymentRequest(AppSettings s) => s.BankName.Length == 0 ? "" : $"Mohon ditransfer ke rekening {s.BankName} :";
    private static string AccountLine(AppSettings s) => string.Join(" - ", new[] { s.BankAccountHolder, s.BankAccountNumber }.Where(x => x.Length > 0));
    private static void RequireRecords<T>(IReadOnlyCollection<T> records) { if (records.Count == 0) throw new InvalidOperationException("Pilih setidaknya satu perjalanan untuk diekspor."); }
    private static void RequireMaximum<T>(IReadOnlyCollection<T> records, int maximum, string layout) { RequireRecords(records); if (records.Count > maximum) throw new InvalidOperationException($"{layout} hanya memuat {maximum} baris. Kurangi pilihan data lalu coba lagi."); }
    private static string IndonesianDate(DateTime d) { string[] m = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]; return $"{d.Day} {m[d.Month - 1]} {d.Year}"; }
    private static string SheetName(string prefix, string number) { var clean = new string(number.Where(char.IsLetterOrDigit).ToArray()); var value = clean.Length == 0 ? prefix : $"{prefix} {clean}"; return value[..Math.Min(value.Length, 31)]; }
    private static string SafeSheetName(string value, XLWorkbook book) { var bad = new HashSet<char>(['[',']',':','*','?','/','\\']); var clean = new string(value.Where(c => !bad.Contains(c)).ToArray()).Trim(); clean = clean.Length == 0 ? "Truk" : clean[..Math.Min(clean.Length,31)]; var candidate = clean; for (var n = 2; book.Worksheets.Any(s => s.Name.Equals(candidate,StringComparison.OrdinalIgnoreCase)); n++) { var end = $" {n}"; candidate = clean[..Math.Min(clean.Length,31-end.Length)] + end; } return candidate; }
    private static void Save(XLWorkbook book, string path) { var dir = Path.GetDirectoryName(path); if (dir?.Length > 0) Directory.CreateDirectory(dir); book.SaveAs(path); }
    private sealed record LedgerEntry(DateTime Date, int KindOrder, long Id, HaulRecord? Haul, ExpenseRecord? Expense);
}
