namespace TruckExcelAssistant.Models;

public sealed record LegacyImportResult(
    int AddedHauls,
    int UpdatedHauls,
    int AddedExpenses,
    int SkippedRows,
    DateTime? LatestDate) : EventArgs;
