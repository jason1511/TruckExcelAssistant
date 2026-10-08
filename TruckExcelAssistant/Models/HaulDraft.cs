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
    decimal ClaimWeightKg,
    decimal ClaimRatePerKg,
    decimal ClaimAmount,
    decimal DriverRoadMoney,
    decimal OtherExpense,
    string Notes,
    OutputLayout Layout)
{
    public decimal WeightDifferenceKg => LoadedWeightKg - ReceivedWeightKg;

    public decimal GrossAmount => ReceivedWeightKg * RatePerKg;

    public decimal CalculatedClaimAmount => ClaimWeightKg * ClaimRatePerKg;

    public decimal EffectiveClaimAmount => CalculatedClaimAmount > 0
        ? CalculatedClaimAmount
        : ClaimAmount;

    public decimal FinalAmount => Layout switch
    {
        OutputLayout.MigunoLike => GrossAmount - BonSangu,
        OutputLayout.AgricoLike => GrossAmount - EffectiveClaimAmount,
        OutputLayout.SumberPanganLike => GrossAmount + RejectionCost,
        _ => GrossAmount - DriverRoadMoney - OtherExpense
    };
}
