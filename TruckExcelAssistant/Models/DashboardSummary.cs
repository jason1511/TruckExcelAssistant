namespace TruckExcelAssistant.Models;

public sealed record DashboardSummary(
    DateTime Month,
    int HaulCount,
    decimal Revenue,
    decimal EmbeddedExpenses,
    decimal StandaloneExpenses,
    decimal Net,
    int OutstandingInvoiceCount,
    decimal OutstandingInvoiceAmount,
    IReadOnlyList<TruckFinancialSummary> Trucks,
    IReadOnlyList<InvoiceRecord> RecentInvoices)
{
    public decimal TotalExpenses => EmbeddedExpenses + StandaloneExpenses;
}
