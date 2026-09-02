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
    OutputLayout Layout)
{
    public decimal WeightDifferenceKg => LoadedWeightKg - ReceivedWeightKg;

    public decimal GrossAmount => ReceivedWeightKg * RatePerKg;

    public decimal FinalAmount => Layout switch
    {
        OutputLayout.CompactInvoice => GrossAmount - BonSangu,
        OutputLayout.CompleteInvoice => GrossAmount + RejectionCost - ClaimAmount,
        _ => GrossAmount - DriverRoadMoney - OtherExpense
    };
}
