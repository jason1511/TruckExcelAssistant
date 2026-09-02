namespace TruckExcelAssistant.Models;

public sealed record HaulRecord(
    long Id,
    HaulDraft Draft,
    HaulStatus Status,
    DateTime CreatedAt,
    DateTime UpdatedAt,
    DateTime? DeletedAt);
