namespace TruckExcelAssistant.Models;

public sealed record TruckFinancialSummary(
    string LicencePlate,
    int HaulCount,
    int ExpenseCount,
    decimal Revenue,
    decimal Expenses,
    decimal Net);
