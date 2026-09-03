using System.Globalization;
using ClosedXML.Excel;
using TruckExcelAssistant.Models;

namespace TruckExcelAssistant.Services;

public sealed class LegacyWorkbookImporter
{
    private const string ImportTag = "[IMPOR EXCEL]";
    private readonly DatabaseService _database;
    private int _addedHauls;
    private int _updatedHauls;
    private int _addedExpenses;
    private int _skippedRows;
    private DateTime? _latestDate;

    public LegacyWorkbookImporter(DatabaseService database)
    {
        _database = database;
    }

    public LegacyImportResult Import(IReadOnlyCollection<string> paths)
    {
        if (paths.Count == 0)
        {
            throw new InvalidOperationException("Pilih setidaknya satu file Excel.");
        }

        var files = paths.Where(File.Exists).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (files.Count != paths.Count)
        {
            throw new FileNotFoundException("Salah satu file Excel yang dipilih tidak ditemukan.");
        }

        foreach (var path in files.Where(IsLedgerFile))
        {
            ImportLedger(path);
        }
        foreach (var path in files.Where(path => !IsLedgerFile(path)))
        {
            ImportInvoiceWorkbook(path);
        }

        return new LegacyImportResult(_addedHauls, _updatedHauls, _addedExpenses, _skippedRows, _latestDate);
    }

    private void ImportLedger(string path)
    {
        using var workbook = new XLWorkbook(path);
        foreach (var sheet in workbook.Worksheets)
        {
            var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
            DateTime? currentDate = null;
            for (var row = 1; row <= lastRow; row++)
            {
                currentDate = ReadDate(sheet.Cell(row, 1)) ?? currentDate;
                var origin = Text(sheet.Cell(row, 2));
                var destination = Text(sheet.Cell(row, 3));
                var cargo = Text(sheet.Cell(row, 4));
                var weight = Number(sheet.Cell(row, 5));
                var rate = Number(sheet.Cell(row, 6));
                var description = Text(sheet.Cell(row, 8));
                var roadMoney = Number(sheet.Cell(row, 9));
                var otherExpense = Number(sheet.Cell(row, 10));
                var sourceKey = SourceKey(path, sheet.Name, row, "ledger");

                if (!string.IsNullOrWhiteSpace(origin)
                    && !string.IsNullOrWhiteSpace(destination)
                    && !string.IsNullOrWhiteSpace(cargo)
                    && weight > 0
                    && rate > 0
                    && currentDate.HasValue)
                {
                    ImportHaul(
                        sourceKey,
                        path,
                        sheet.Name,
                        row,
                        new HaulDraft(
                            currentDate.Value, sheet.Name.Trim(), cargo, string.Empty, origin, destination,
                            weight, weight, rate, 0, 0, 0, roadMoney, otherExpense,
                            $"{ImportTag} {Path.GetFileName(path)} / {sheet.Name} / baris {row}",
                            OutputLayout.TruckLedger),
                        preferIncomingInvoiceFields: false);
                    continue;
                }

                if (currentDate.HasValue && !string.IsNullOrWhiteSpace(description) && otherExpense > 0)
                {
                    if (_database.HasLegacyImportRow(sourceKey))
                    {
                        _skippedRows++;
                        continue;
                    }
                    var id = _database.AddExpense(
                        currentDate.Value,
                        sheet.Name.Trim(),
                        ExpenseCategory(description),
                        $"{ImportTag} {description}",
                        otherExpense);
                    _database.RecordLegacyImportRow(sourceKey, path, sheet.Name, row, "Expense", id);
                    _addedExpenses++;
                    TrackDate(currentDate.Value);
                }
            }
        }
    }

    private void ImportInvoiceWorkbook(string path)
    {
        using var workbook = new XLWorkbook(path);
        var claims = ReadClaims(workbook);
        foreach (var sheet in workbook.Worksheets.Where(item => item.Name.StartsWith("INV", StringComparison.OrdinalIgnoreCase)))
        {
            var headerRow = FindHeaderRow(sheet);
            if (headerRow == 0)
            {
                continue;
            }
            var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            var dateColumn = FindColumn(sheet, headerRow, lastColumn, value => value is "TGL" or "TANGGAL");
            var cargoColumn = FindColumn(sheet, headerRow, lastColumn, value => value.StartsWith("JENIS", StringComparison.Ordinal));
            var plateColumn = FindColumn(sheet, headerRow, lastColumn, value => value == "NOPOL");
            var weightColumns = FindColumns(sheet, headerRow, lastColumn, value => value.StartsWith("BERAT", StringComparison.Ordinal));
            var rateColumn = FindColumn(sheet, headerRow, lastColumn, value => value is "ONGK" or "ONGKOS");
            var bonColumn = FindColumn(sheet, headerRow, lastColumn, value => value.Contains("BON SANGU", StringComparison.Ordinal));
            var rejectionColumn = FindColumn(sheet, headerRow, lastColumn, value => value.StartsWith("BIAYA TOLAKAN", StringComparison.Ordinal));
            if (dateColumn == 0 || cargoColumn == 0 || plateColumn == 0 || weightColumns.Count == 0 || rateColumn == 0)
            {
                continue;
            }

            var customer = CustomerName(sheet.Cell(1, 1).GetFormattedString());
            var compact = bonColumn > 0;
            var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
            for (var row = headerRow + 1; row <= lastRow; row++)
            {
                if (Number(sheet.Cell(row, 1)) <= 0)
                {
                    continue;
                }
                var date = ReadDate(sheet.Cell(row, dateColumn));
                var plate = Text(sheet.Cell(row, plateColumn));
                var cargo = Text(sheet.Cell(row, cargoColumn));
                var rate = Number(sheet.Cell(row, rateColumn));
                var loaded = Number(sheet.Cell(row, weightColumns[0]));
                var received = weightColumns.Count > 1 ? Number(sheet.Cell(row, weightColumns[1])) : loaded;
                if (!date.HasValue || string.IsNullOrWhiteSpace(plate) || string.IsNullOrWhiteSpace(cargo) || received <= 0 || rate <= 0)
                {
                    continue;
                }
                claims.TryGetValue(ClaimKey(date.Value, plate, loaded, received), out var claim);
                var draft = new HaulDraft(
                    date.Value, plate, cargo, customer, string.Empty, string.Empty,
                    loaded, received, rate,
                    bonColumn > 0 ? Number(sheet.Cell(row, bonColumn)) : 0,
                    rejectionColumn > 0 ? Number(sheet.Cell(row, rejectionColumn)) : 0,
                    claim,
                    0, 0,
                    $"{ImportTag} {Path.GetFileName(path)} / {sheet.Name} / baris {row}",
                    compact ? OutputLayout.CompactInvoice : OutputLayout.CompleteInvoice);
                ImportHaul(SourceKey(path, sheet.Name, row, "invoice"), path, sheet.Name, row, draft, preferIncomingInvoiceFields: true);
            }
        }
    }

    private void ImportHaul(
        string sourceKey,
        string path,
        string sheet,
        int row,
        HaulDraft incoming,
        bool preferIncomingInvoiceFields)
    {
        if (_database.HasLegacyImportRow(sourceKey))
        {
            _skippedRows++;
            return;
        }

        var existing = _database.FindMatchingHaul(incoming.Date, incoming.LicencePlate, incoming.Cargo, incoming.ReceivedWeightKg, incoming.RatePerKg);
        long id;
        if (existing is null)
        {
            id = _database.AddHaul(incoming, HaulStatus.Saved);
            _addedHauls++;
        }
        else
        {
            var old = existing.Draft;
            var merged = preferIncomingInvoiceFields
                ? incoming with
                {
                    Origin = old.Origin,
                    Destination = old.Destination,
                    DriverRoadMoney = old.DriverRoadMoney,
                    OtherExpense = old.OtherExpense,
                    Notes = CombineNotes(old.Notes, incoming.Notes)
                }
                : old with
                {
                    Origin = incoming.Origin,
                    Destination = incoming.Destination,
                    DriverRoadMoney = incoming.DriverRoadMoney,
                    OtherExpense = incoming.OtherExpense,
                    Notes = CombineNotes(old.Notes, incoming.Notes)
                };
            _database.UpdateHaul(existing.Id, merged, HaulStatus.Saved);
            id = existing.Id;
            _updatedHauls++;
        }
        _database.RecordLegacyImportRow(sourceKey, path, sheet, row, "Haul", id);
        TrackDate(incoming.Date);
    }

    private static Dictionary<string, decimal> ReadClaims(XLWorkbook workbook)
    {
        var claims = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in workbook.Worksheets.Where(item => item.Name.StartsWith("KLAIM", StringComparison.OrdinalIgnoreCase)))
        {
            var headerRow = FindHeaderRow(sheet);
            var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            if (headerRow == 0 || lastColumn == 0)
            {
                continue;
            }
            var dateColumn = FindColumn(sheet, headerRow, lastColumn, value => value is "TGL" or "TANGGAL");
            var plateColumn = FindColumn(sheet, headerRow, lastColumn, value => value == "NOPOL");
            var weightColumns = FindColumns(sheet, headerRow, lastColumn, value => value.StartsWith("BERAT", StringComparison.Ordinal));
            var amountColumn = FindColumn(sheet, headerRow, lastColumn, value => value == "JUMLAH");
            if (dateColumn == 0 || plateColumn == 0 || weightColumns.Count < 2 || amountColumn == 0)
            {
                continue;
            }
            var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
            for (var row = headerRow + 1; row <= lastRow; row++)
            {
                if (Number(sheet.Cell(row, 1)) <= 0)
                {
                    continue;
                }
                var date = ReadDate(sheet.Cell(row, dateColumn));
                if (date.HasValue)
                {
                    claims[ClaimKey(
                        date.Value,
                        Text(sheet.Cell(row, plateColumn)),
                        Number(sheet.Cell(row, weightColumns[0])),
                        Number(sheet.Cell(row, weightColumns[1])))] = Number(sheet.Cell(row, amountColumn));
                }
            }
        }
        return claims;
    }

    private static int FindHeaderRow(IXLWorksheet sheet)
    {
        var lastColumn = Math.Min(sheet.LastColumnUsed()?.ColumnNumber() ?? 0, 20);
        for (var row = 1; row <= Math.Min(sheet.LastRowUsed()?.RowNumber() ?? 0, 12); row++)
        {
            for (var column = 1; column <= lastColumn; column++)
            {
                if (Header(sheet.Cell(row, column)) == "NO")
                {
                    return row;
                }
            }
        }
        return 0;
    }

    private static int FindColumn(IXLWorksheet sheet, int row, int lastColumn, Func<string, bool> predicate) =>
        FindColumns(sheet, row, lastColumn, predicate).FirstOrDefault();

    private static List<int> FindColumns(IXLWorksheet sheet, int row, int lastColumn, Func<string, bool> predicate)
    {
        var columns = new List<int>();
        for (var column = 1; column <= lastColumn; column++)
        {
            if (predicate(Header(sheet.Cell(row, column))))
            {
                columns.Add(column);
            }
        }
        return columns;
    }

    private static string Header(IXLCell cell) => new(cell.GetFormattedString()
        .Trim()
        .ToUpperInvariant()
        .Where(character => char.IsLetterOrDigit(character) || char.IsWhiteSpace(character))
        .ToArray());

    private static string Text(IXLCell cell) => cell.GetFormattedString().Trim();

    private static decimal Number(IXLCell cell)
    {
        if (cell.TryGetValue<double>(out var numeric))
        {
            return Convert.ToDecimal(numeric);
        }
        var text = cell.GetFormattedString().Trim();
        return decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var value)
            || decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value)
            ? value
            : 0;
    }

    private static DateTime? ReadDate(IXLCell cell)
    {
        if (cell.TryGetValue<DateTime>(out var date))
        {
            return date.Date;
        }
        if (cell.TryGetValue<double>(out var serial) && serial is >= 30_000 and <= 60_000)
        {
            return DateTime.FromOADate(serial).Date;
        }
        return null;
    }

    private static bool IsLedgerFile(string path) =>
        Path.GetFileName(path).Contains("PEMBUKUAN", StringComparison.OrdinalIgnoreCase);

    private static string CustomerName(string heading)
    {
        var value = heading.Replace("Kepada Yth.", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(":", string.Empty)
            .Trim();
        return value;
    }

    private static string CombineNotes(string first, string second) =>
        string.IsNullOrWhiteSpace(first) ? second
        : string.IsNullOrWhiteSpace(second) || first.Contains(second, StringComparison.OrdinalIgnoreCase) ? first
        : $"{first}; {second}";

    private static string ExpenseCategory(string description)
    {
        var value = description.ToUpperInvariant();
        if (value.Contains("BAN") || value.Contains("PELEG")) return "Ban";
        if (value.Contains("SOLAR")) return "Bahan bakar";
        if (value.Contains("STNK") || value.Contains("PAJAK")) return "Pajak & administrasi";
        if (value.Contains("SERVICE") || value.Contains("SERVIS") || value.Contains("OLI") || value.Contains("LAKER")) return "Servis & suku cadang";
        if (value.Contains("TOL") || value.Contains("PARKIR")) return "Tol & parkir";
        return "Lainnya";
    }

    private static string SourceKey(string path, string sheet, int row, string kind) =>
        $"{Path.GetFileName(path).ToUpperInvariant()}|{sheet.ToUpperInvariant()}|{row}|{kind}";

    private static string ClaimKey(DateTime date, string plate, decimal loaded, decimal received) =>
        $"{date:yyyyMMdd}|{plate.Trim().ToUpperInvariant()}|{loaded:0.###}|{received:0.###}";

    private void TrackDate(DateTime date)
    {
        if (!_latestDate.HasValue || date > _latestDate.Value)
        {
            _latestDate = date;
        }
    }
}
