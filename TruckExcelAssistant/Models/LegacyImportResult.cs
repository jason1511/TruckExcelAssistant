namespace TruckExcelAssistant.Models;

public sealed class LegacyImportResult : EventArgs
{
    public LegacyImportResult(
        int addedHauls,
        int updatedHauls,
        int addedExpenses,
        int skippedRows,
        DateTime? latestDate)
    {
        AddedHauls = addedHauls;
        UpdatedHauls = updatedHauls;
        AddedExpenses = addedExpenses;
        SkippedRows = skippedRows;
        LatestDate = latestDate;
    }

    public int AddedHauls { get; }
    public int UpdatedHauls { get; }
    public int AddedExpenses { get; }
    public int SkippedRows { get; }
    public DateTime? LatestDate { get; }
}
