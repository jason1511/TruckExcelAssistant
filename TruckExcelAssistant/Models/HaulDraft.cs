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
    OutputMode Mode)
{
    public decimal WeightDifferenceKg => LoadedWeightKg - ReceivedWeightKg;

    public decimal GrossAmount => ReceivedWeightKg * RatePerKg;

    public decimal FinalAmount => Mode switch
    {
        OutputMode.Miguno => GrossAmount - BonSangu,
        OutputMode.Agrico => GrossAmount + RejectionCost - ClaimAmount,
        _ => GrossAmount - DriverRoadMoney - OtherExpense
    };
}
