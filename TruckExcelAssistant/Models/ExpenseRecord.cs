namespace TruckExcelAssistant.Models;

public sealed record ExpenseRecord(
    long Id,
    DateTime Date,
    string LicencePlate,
    string Category,
    string Description,
    decimal Amount,
    DateTime CreatedAt,
    DateTime UpdatedAt,
    DateTime? DeletedAt);
