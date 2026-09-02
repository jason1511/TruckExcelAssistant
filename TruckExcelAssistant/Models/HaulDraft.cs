namespace TruckExcelAssistant.Models;

public sealed record HaulDraft(
    DateTime Date,
    string LicencePlate,
    string Cargo,
    string Customer,
    string Origin,
    string Destination,
    decimal LoadedWeightKg,
    decimal ReceivedWeightKg,
    decimal RatePerKg,
    decimal BonSangu,
    decimal RejectionCost,
    decimal ClaimAmount,
    decimal DriverRoadMoney,
    decimal OtherExpense,
    string Notes,
    CustomerKind CustomerType)
{
    public decimal WeightDifferenceKg => LoadedWeightKg - ReceivedWeightKg;

    public decimal GrossAmount => ReceivedWeightKg * RatePerKg;

    public decimal InvoiceAmount => CustomerType switch
    {
        CustomerKind.Miguno => GrossAmount - BonSangu,
        CustomerKind.Agrico => GrossAmount + RejectionCost - ClaimAmount,
        _ => GrossAmount + RejectionCost - BonSangu - ClaimAmount
    };

    public decimal LedgerNetAmount => GrossAmount - DriverRoadMoney - OtherExpense;
}
