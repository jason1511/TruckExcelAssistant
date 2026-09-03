namespace TruckExcelAssistant.Models;

public sealed record InvoiceRecord(
    long Id,
    string InvoiceNumber,
    DateTime InvoiceDate,
    string Customer,
    OutputLayout Layout,
    decimal TotalAmount,
    string FilePath,
    InvoiceStatus Status,
    DateTime CreatedAt,
    int HaulCount);
