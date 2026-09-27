using System;
using System.Globalization;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public enum ForecastTargetMode { DefaultP80, Manual, None }

// Detached editor state: opening, cancelling and invalid input never mutate the model.
public sealed class ForecastEditDraft
{
    private readonly ForecastTargetSettings? original;
    public string Name { get; set; }
    public ForecastTargetMode Mode { get; set; }
    public string TargetText { get; set; }
    public TargetDirection? Direction { get; set; }
    public double? RequestedConfidence => original?.RequestedConfidence;

    public ForecastEditDraft(ForecastDefinition forecast)
    {
        Name = forecast.Name;
        original = forecast.TargetSettings;
        Mode = original == null ? ForecastTargetMode.DefaultP80 :
            original.Target.HasValue ? ForecastTargetMode.Manual : ForecastTargetMode.None;
        TargetText = original?.Target?.ToString("R", CultureInfo.CurrentCulture) ?? "";
        Direction = original?.Direction;
    }

    public bool TryApply(ForecastDefinition forecast, out string? error)
    {
        error = null;
        if (!Enum.IsDefined(Mode) || (Direction.HasValue && !Enum.IsDefined(Direction.Value)))
        {
            error = "Select a valid target mode and direction.";
            return false;
        }
        double? target = null;
        if (Mode == ForecastTargetMode.Manual)
        {
            if (!double.TryParse(TargetText, NumberStyles.Float | NumberStyles.AllowThousands,
                    CultureInfo.CurrentCulture, out double value) || !double.IsFinite(value))
            {
                error = "Enter a finite numeric target.";
                return false;
            }
            target = value;
        }
        // A confidence-derived target retains its provenance only while target and direction agree.
        var confidence = target.HasValue && target == original?.Target && Direction == original?.Direction
            ? original?.RequestedConfidence : null;
        var settings = Mode == ForecastTargetMode.DefaultP80 ? null :
            new ForecastTargetSettings(target, Direction, confidence);
        forecast.Name = string.IsNullOrWhiteSpace(Name) ? $"{forecast.SheetName}!{forecast.CellAddress}" : Name.Trim();
        forecast.TargetSettings = settings;
        return true;
    }
}
