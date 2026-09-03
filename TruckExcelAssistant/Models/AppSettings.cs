namespace TruckExcelAssistant.Models;

public sealed record AppSettings(
    string CompanyName,
    string CompanyAddress,
    string City,
    string BankName,
    string BankAccountNumber,
    string BankAccountHolder,
    string SignerName,
    string InvoicePrefix,
    int InvoiceSequenceDigits,
    OutputLayout DefaultInvoiceLayout,
    string DefaultExportDirectory)
{
    public static AppSettings Default => new(
        string.Empty,
        string.Empty,
        "Lumajang",
        string.Empty,
        string.Empty,
        string.Empty,
        string.Empty,
        "TJ",
        3,
        OutputLayout.CompleteInvoice,
        string.Empty);
}
