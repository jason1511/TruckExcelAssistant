namespace TruckExcelAssistant.Models;

public enum OutputLayout
{
    TruckLedger = 0,
    MigunoLike = 1,
    SumberPanganLike = 2,
    AgricoLike = 3
}

public static class OutputLayoutNames
{
    public static string DisplayName(this OutputLayout layout) => layout switch
    {
        OutputLayout.MigunoLike => "Mirip Miguno",
        OutputLayout.AgricoLike => "Mirip Agrico",
        OutputLayout.SumberPanganLike => "Mirip Sumber Pangan",
        _ => "Pembukuan Truk"
    };

    public static OutputLayout ParseInvoiceSetting(string? value, OutputLayout fallback)
    {
        if (string.Equals(value, "CompactInvoice", StringComparison.OrdinalIgnoreCase))
        {
            return OutputLayout.MigunoLike;
        }
        if (string.Equals(value, "CompleteInvoice", StringComparison.OrdinalIgnoreCase))
        {
            return OutputLayout.SumberPanganLike;
        }
        return Enum.TryParse<OutputLayout>(value, true, out var parsed)
               && parsed is OutputLayout.MigunoLike or OutputLayout.AgricoLike or OutputLayout.SumberPanganLike
            ? parsed
            : fallback;
    }
}
