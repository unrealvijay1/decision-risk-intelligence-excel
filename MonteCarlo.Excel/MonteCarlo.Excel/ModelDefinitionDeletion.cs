using System;
using System.Linq;
using ExcelDna.Integration;

namespace MonteCarlo.Excel;

public enum DeleteDefinitionResult { Deleted, Cancelled, NoDefinition, InvalidSelection }

public static class ModelDefinitionDeletion
{
    public const string SelectionMessage = "Please select a single Assumption or Forecast cell.";
    public const string NoDefinitionMessage = "The selected cell does not contain a Monte Carlo assumption or forecast.";

    // Shared by Model Manager and the ribbon. No mutation occurs before confirmation.
    public static DeleteDefinitionResult Delete(object definition, Func<string, string, bool> confirm)
    {
        string kind, name, sheet, address;
        if (definition is AssumptionDefinition a && SimulationModel.Assumptions.Contains(a))
            (kind, name, sheet, address) = ("Assumption", a.Name, a.SheetName, a.CellAddress);
        else if (definition is ForecastDefinition f && SimulationModel.Forecasts.Contains(f))
            (kind, name, sheet, address) = ("Forecast", f.Name, f.SheetName, f.CellAddress);
        else return DeleteDefinitionResult.NoDefinition;
        if (!confirm($"Remove {kind}?", $"Remove the {kind} definition '{name}' from {sheet}!{address}?\n\nThe Excel cell value/formula will not be changed."))
            return DeleteDefinitionResult.Cancelled;
        if (definition is AssumptionDefinition assumption) SimulationModel.Assumptions.Remove(assumption);
        else SimulationModel.Forecasts.Remove((ForecastDefinition)definition);
        WorkbookPersistence.SaveModel();
        return DeleteDefinitionResult.Deleted;
    }

    public static DeleteDefinitionResult ClearSelectedCell(object? selection, Func<string, string, bool> confirm)
    {
        string sheet, address;
        try
        {
            if (selection == null) return DeleteDefinitionResult.InvalidSelection;
            dynamic cell = selection;
            dynamic app = ExcelDnaUtil.Application;
            if (Convert.ToDouble(cell.Cells.CountLarge) != 1 || Convert.ToInt32(cell.Areas.Count) != 1 ||
                !string.Equals((string)cell.Worksheet.Parent.Name, (string)app.ActiveWorkbook.Name, StringComparison.OrdinalIgnoreCase))
                return DeleteDefinitionResult.InvalidSelection;
            sheet = cell.Worksheet.Name;
            address = cell.Address[false, false];
        }
        catch (Exception ex) when (ex is System.Runtime.InteropServices.COMException || ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        {
            return DeleteDefinitionResult.InvalidSelection;
        }
        var assumptions = SimulationModel.Assumptions.ToArray();
        var forecasts = SimulationModel.Forecasts.ToArray();
        bool deleted = false;
        try
        {
            // Resolve tracked names without migrating fill metadata or touching cell interiors.
            WorkbookPersistence.LoadModelForSimulation(restoreHighlights: false);
            object? definition = (object?)SimulationModel.FindAssumption(sheet, address) ?? SimulationModel.FindForecast(sheet, address);
            var result = definition == null ? DeleteDefinitionResult.NoDefinition : Delete(definition, confirm);
            deleted = result == DeleteDefinitionResult.Deleted;
            return result;
        }
        finally
        {
            if (!deleted)
            {
                SimulationModel.Clear();
                SimulationModel.Assumptions.AddRange(assumptions);
                SimulationModel.Forecasts.AddRange(forecasts);
            }
        }
    }
}
