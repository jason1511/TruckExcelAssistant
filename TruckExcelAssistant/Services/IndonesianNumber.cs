using System.Globalization;

namespace TruckExcelAssistant.Services;

internal static class IndonesianNumber
{
    private static readonly CultureInfo Culture = CultureInfo.GetCultureInfo("id-ID");

    public static string Format(decimal value) => value.ToString("N0", Culture);

    public static string Rupiah(decimal value) => $"Rp{Format(value)}";

    public static bool TryParse(string? text, out decimal value)
    {
        var normalized = (text ?? string.Empty)
            .Replace("Rp", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(" ", string.Empty)
            .Replace(".", string.Empty)
            .Replace(",", string.Empty);

        return decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out value);
    }
}
