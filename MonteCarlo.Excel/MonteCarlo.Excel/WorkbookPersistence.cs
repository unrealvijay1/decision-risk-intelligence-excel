using System;
using MonteCarlo.Core;
using ExcelDna.Integration;

namespace MonteCarlo.Excel
{
    public static class WorkbookPersistence
    {
        private const string ConfigSheetName =
            "__MonteCarloConfig";


        // =========================================================
        // SAVE MODEL TO WORKBOOK
        // =========================================================

        public static void SaveModel()
        {
            dynamic excelApp =
                ExcelDnaUtil.Application;


            dynamic workbook =
                excelApp.ActiveWorkbook;


            if (workbook == null)
            {
                throw new Exception(
                    "No active workbook.");
            }


            dynamic configSheet =
                GetOrCreateConfigSheet(
                    workbook);


            // Resolve/create links before changing any existing metadata.
            foreach (var assumption in SimulationModel.Assumptions)
                assumption.CellLink = EnsureCellLink(workbook, assumption.CellLink, assumption.SheetName, assumption.CellAddress);
            foreach (var forecast in SimulationModel.Forecasts)
                forecast.CellLink = EnsureCellLink(workbook, forecast.CellLink, forecast.SheetName, forecast.CellAddress);
            configSheet.Cells.Clear();
            // Text metadata must never become worksheet formulas.
            configSheet.Range["A:D"].NumberFormat = "@";
            configSheet.Range["I:N"].NumberFormat = "@";


            // =====================================================
            // HEADER
            // =====================================================

            configSheet.Cells[1, 1].Value2 =
                "Type";

            configSheet.Cells[1, 2].Value2 =
                "Sheet";

            configSheet.Cells[1, 3].Value2 =
                "Cell";

            configSheet.Cells[1, 4].Value2 =
                "Distribution";

            configSheet.Cells[1, 5].Value2 =
                "Parameter1";

            configSheet.Cells[1, 6].Value2 =
                "Parameter2";

            configSheet.Cells[1, 7].Value2 =
                "Parameter3";

            configSheet.Cells[1, 8].Value2 =
                "Parameter4";

            configSheet.Cells[1, 9].Value2 =
                "Name";


            configSheet.Cells[1, 10].Value2 = "CellLink";

            int row =
                2;



            // =====================================================
            // SAVE ASSUMPTIONS
            // =====================================================

            foreach (
                AssumptionDefinition assumption
                in SimulationModel.Assumptions)
            {
                configSheet.Cells[
                    row,
                    1].Value2 =
                    "Assumption";


                configSheet.Cells[
                    row,
                    2].Value2 =
                    assumption.SheetName;


                configSheet.Cells[
                    row,
                    3].Value2 =
                    assumption.CellAddress;


                configSheet.Cells[
                    row,
                    4].Value2 =
                    assumption
                        .Distribution
                        .ToString();


                configSheet.Cells[
                    row,
                    5].Value2 =
                    assumption.Parameter1;


                configSheet.Cells[
                    row,
                    6].Value2 =
                    assumption.Parameter2;


                configSheet.Cells[
                    row,
                    7].Value2 =
                    assumption.Parameter3;


                configSheet.Cells[
                    row,
                    8].Value2 =
                    assumption.Parameter4;


                configSheet.Cells[
                    row,
                    9].Value2 =
                    assumption.Name;


                configSheet.Cells[row, 10].Value2 = assumption.CellLink;
                row++;
            }


            // =====================================================
            // SAVE FORECASTS
            // =====================================================

            foreach (
                ForecastDefinition forecast
                in SimulationModel.Forecasts)
            {
                configSheet.Cells[
                    row,
                    1].Value2 =
                    "Forecast";


                configSheet.Cells[
                    row,
                    2].Value2 =
                    forecast.SheetName;


                configSheet.Cells[
                    row,
                    3].Value2 =
                    forecast.CellAddress;


                configSheet.Cells[
                    row,
                    4].Value2 =
                    "";


                configSheet.Cells[
                    row,
                    5].Value2 =
                    "";


                configSheet.Cells[
                    row,
                    6].Value2 =
                    "";


                configSheet.Cells[
                    row,
                    7].Value2 =
                    "";


                configSheet.Cells[
                    row,
                    8].Value2 =
                    "";


                configSheet.Cells[
                    row,
                    9].Value2 =
                    forecast.Name;


                configSheet.Cells[row, 10].Value2 = forecast.CellLink;
                WriteTargetSettings(configSheet, row, forecast.TargetSettings);
                row++;
            }


            // =====================================================
            // VERY HIDDEN CONFIG SHEET
            // =====================================================

            configSheet.Visible =
                2;
        }


        // =========================================================
        // LOAD MODEL FROM WORKBOOK
        // =========================================================

        public static ValidationResult LastLoadValidation { get; private set; } = new();
        public static Action<string>? ReportRestoreProblem { get; set; }

        public static void LoadModel()
        {
            LastLoadValidation = LoadModelForSimulation();
            if (!LastLoadValidation.IsValid) ReportRestoreProblem?.Invoke(LastLoadValidation.UserMessage);
        }

        public static ValidationResult LoadModelForSimulation()
        {
            var validation = new ValidationResult();
            LoadModelCore(validation);
            return validation;
        }

        private static void LoadModelCore(ValidationResult? validation)
        {
            dynamic excelApp =
                ExcelDnaUtil.Application;


            dynamic workbook =
                excelApp.ActiveWorkbook;


            SimulationModel.Clear();


            if (workbook == null)
            {
                return;
            }


            dynamic configSheet =
                FindConfigSheet(
                    workbook);


            if (configSheet == null)
            {
                return;
            }


            bool newLayout =
                IsNewLayout(
                    configSheet);


            int row =
                2;


            // Simulation validation must not silently truncate a damaged saved model at a blank row.
            int lastConfigRow = validation == null ? 0 :
                (int)configSheet.UsedRange.Row + (int)configSheet.UsedRange.Rows.Count - 1;

            while (true)
            {
                object typeValue =
                    configSheet.Cells[
                        row,
                        1].Value2;


                if (typeValue == null)
                {
                    if (validation == null || row > lastConfigRow) break;
                    bool hasData = false;
                    for (int column = 2; column <= (newLayout ? 14 : 8); column++)
                        if (configSheet.Cells[row, column].Value2 != null) hasData = true;
                    if (hasData)
                        validation.Add($"Saved model row {row}", "The item type is missing.",
                            "Use Model Manager to redefine this saved assumption or forecast.");
                    row++;
                    continue;
                }


                string type =
                    Convert.ToString(
                        typeValue)
                    ?? "";


                if (
                    string.Equals(
                        type,
                        "Assumption",
                        StringComparison.OrdinalIgnoreCase))
                {
                    LoadAssumption(
                        configSheet,
                        row,
                        newLayout, validation);
                }


                else if (
                    string.Equals(
                        type,
                        "Forecast",
                        StringComparison.OrdinalIgnoreCase))
                {
                    LoadForecast(
                        configSheet,
                        row,
                        newLayout, validation);
                }


                else if (validation != null)
                    validation.Add($"Saved model row {row}", "The item type is not recognized.",
                        "Use Model Manager to remove or redefine this saved item.");
                row++;
            }
        }


        // =========================================================
        // DETECT CONFIG VERSION
        //
        // OLD:
        // 8 = Name
        //
        // NEW:
        // 8 = Parameter4
        // 9 = Name
        // =========================================================

        private static bool IsNewLayout(
            dynamic configSheet)
        {
            string column8Header =
                Convert.ToString(
                    configSheet.Cells[
                        1,
                        8].Value2)
                ?? "";


            string column9Header =
                Convert.ToString(
                    configSheet.Cells[
                        1,
                        9].Value2)
                ?? "";


            return
                string.Equals(
                    column8Header,
                    "Parameter4",
                    StringComparison.OrdinalIgnoreCase)
                ||
                string.Equals(
                    column9Header,
                    "Name",
                    StringComparison.OrdinalIgnoreCase);
        }


        // =========================================================
        // CLEAR SAVED MODEL
        // =========================================================

        public static void ClearSavedModel()
        {
            dynamic excelApp =
                ExcelDnaUtil.Application;


            dynamic workbook =
                excelApp.ActiveWorkbook;


            SimulationModel.Clear();


            if (workbook == null)
            {
                return;
            }


            dynamic configSheet =
                FindConfigSheet(
                    workbook);


            if (configSheet == null)
            {
                return;
            }


            bool originalDisplayAlerts =
                excelApp.DisplayAlerts;


            try
            {
                excelApp.DisplayAlerts =
                    false;


                configSheet.Delete();
            }
            finally
            {
                excelApp.DisplayAlerts =
                    originalDisplayAlerts;
            }
        }


        // =========================================================
        // LOAD ASSUMPTION
        // =========================================================

        private static void LoadAssumption(
            dynamic configSheet,
            int row,
            bool newLayout, ValidationResult? validation)
        {
            string sheetName =
                Convert.ToString(
                    configSheet.Cells[
                        row,
                        2].Value2)
                ?? "";


            string cellAddress =
                Convert.ToString(
                    configSheet.Cells[
                        row,
                        3].Value2)
                ?? "";


            string distributionText =
                Convert.ToString(
                    configSheet.Cells[
                        row,
                        4].Value2)
                ?? "";


            double parameter1 =
                GetDoubleValue(
                    configSheet.Cells[
                        row,
                        5].Value2);


            double parameter2 =
                GetDoubleValue(
                    configSheet.Cells[
                        row,
                        6].Value2);


            double parameter3 =
                GetDoubleValue(
                    configSheet.Cells[
                        row,
                        7].Value2);


            double parameter4;


            string name;


            if (newLayout)
            {
                parameter4 =
                    GetDoubleValue(
                        configSheet.Cells[
                            row,
                            8].Value2);


                name =
                    Convert.ToString(
                        configSheet.Cells[
                            row,
                            9].Value2)
                    ?? "";
            }
            else
            {
                // ---------------------------------------------
                // OLD WORKBOOK FORMAT
                //
                // Column 8 contained Name.
                // There was no Parameter4.
                // ---------------------------------------------

                parameter4 =
                    0;


                name =
                    Convert.ToString(
                        configSheet.Cells[
                            row,
                            8].Value2)
                    ?? "";
            }


            string cellLink = Convert.ToString(ReadOptional(configSheet, row, 10, "CellLink")) ?? "";
            ResolveSavedLink(configSheet.Parent, cellLink, ref sheetName, ref cellAddress,
                validation, $"Saved assumption row {row} ({sheetName}!{cellAddress})");

            if (validation != null)
            {
                string location = $"Saved assumption row {row} ({sheetName}!{cellAddress})";
                DistributionType savedDistribution = Enum.TryParse(distributionText, true, out DistributionType parsed)
                    ? parsed : (DistributionType)(-1);
                object?[] raw = new object?[4];
                for (int i = 0; i < (newLayout ? 4 : 3); i++)
                    raw[i] = ExcelSimulationWorkbook.ReadCellValue(ExcelDnaUtil.Application, configSheet.Cells[row, 5 + i]);
                SimulationValidation.ValidateParameters(validation, location, savedDistribution, raw);
                if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(cellAddress))
                    validation.Add(location, "The worksheet or cell reference is missing.", "Redefine the assumption using a single input cell.");
            }

            if (
                string.IsNullOrWhiteSpace(
                    sheetName)
                ||
                string.IsNullOrWhiteSpace(
                    cellAddress))
            {
                return;
            }


            if (!Enum.TryParse(
                    distributionText,
                    true,
                    out DistributionType distribution))
            {
                return;
            }


            if (string.IsNullOrWhiteSpace(
                    name))
            {
                name =
                    $"{sheetName}!{cellAddress}";
            }


            SimulationModel.Assumptions.Add(
                new AssumptionDefinition
                {
                    CellLink = cellLink,
                    Name =
                        name,

                    SheetName =
                        sheetName,

                    CellAddress =
                        cellAddress,

                    Distribution =
                        distribution,

                    Parameter1 =
                        parameter1,

                    Parameter2 =
                        parameter2,

                    Parameter3 =
                        parameter3,

                    Parameter4 =
                        parameter4
                });
        }


        // =========================================================
        // LOAD FORECAST
        // =========================================================

        private static void LoadForecast(
            dynamic configSheet,
            int row,
            bool newLayout, ValidationResult? validation)
        {
            string sheetName =
                Convert.ToString(
                    configSheet.Cells[
                        row,
                        2].Value2)
                ?? "";


            string cellAddress =
                Convert.ToString(
                    configSheet.Cells[
                        row,
                        3].Value2)
                ?? "";


            string name;


            if (newLayout)
            {
                name =
                    Convert.ToString(
                        configSheet.Cells[
                            row,
                            9].Value2)
                    ?? "";
            }
            else
            {
                name =
                    Convert.ToString(
                        configSheet.Cells[
                            row,
                            8].Value2)
                    ?? "";
            }


            string cellLink = Convert.ToString(ReadOptional(configSheet, row, 10, "CellLink")) ?? "";
            ResolveSavedLink(configSheet.Parent, cellLink, ref sheetName, ref cellAddress,
                validation, $"Saved forecast row {row} ({sheetName}!{cellAddress})");

            if (validation != null && (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(cellAddress)))
                validation.Add($"Saved forecast row {row}", "The worksheet or output cell reference is missing.",
                    "Select the output cell and choose Define Forecast.");

            if (
                string.IsNullOrWhiteSpace(
                    sheetName)
                ||
                string.IsNullOrWhiteSpace(
                    cellAddress))
            {
                return;
            }


            if (string.IsNullOrWhiteSpace(
                    name))
            {
                name =
                    $"{sheetName}!{cellAddress}";
            }


            SimulationModel.Forecasts.Add(
                new ForecastDefinition
                {
                    TargetSettings = ReadTargetSettings(configSheet, row, validation),
                    CellLink = cellLink,
                    Name =
                        name,

                    SheetName =
                        sheetName,

                    CellAddress =
                        cellAddress
                });
        }


        // =========================================================
        // GET OR CREATE CONFIG SHEET
        // =========================================================

        // Update only the forecast row, retaining unrelated (including damaged) saved metadata.
        public static void SaveTargetSettings(ForecastDefinition forecast, ForecastTargetSettings settings, object? sourceWorkbook = null)
        {
            dynamic application = ExcelDnaUtil.Application;
            dynamic workbook = sourceWorkbook ?? application.ActiveWorkbook;
            dynamic? sheet = workbook == null ? null : FindConfigSheet(workbook);
            if (sheet == null) throw new InvalidOperationException("The saved forecast is unavailable.");
            int last = (int)sheet.UsedRange.Row + (int)sheet.UsedRange.Rows.Count - 1;
            for (int row = 2; row <= last; row++)
            {
                if (!string.Equals(Convert.ToString(sheet.Cells[row, 1].Value2), "Forecast", StringComparison.OrdinalIgnoreCase)) continue;
                string link = Convert.ToString(ReadOptional(sheet, row, 10, "CellLink")) ?? "";
                bool matches = !string.IsNullOrEmpty(forecast.CellLink) ? link == forecast.CellLink :
                    string.Equals(Convert.ToString(sheet.Cells[row, 2].Value2), forecast.SheetName, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(Convert.ToString(sheet.Cells[row, 3].Value2), forecast.CellAddress, StringComparison.OrdinalIgnoreCase);
                if (!matches) continue;
                WriteTargetSettings(sheet, row, settings);
                forecast.TargetSettings = settings;
                return;
            }
            throw new InvalidOperationException("The saved forecast is unavailable.");
        }

        private static void WriteTargetSettings(dynamic sheet, int row, ForecastTargetSettings? settings)
        {
            string[] headers = { "TargetConfigured", "Target", "TargetDirection", "RequestedConfidence" };
            for (int i = 0; i < headers.Length; i++) sheet.Cells[1, 11 + i].Value2 = headers[i];
            sheet.Cells[row, 11].Value2 = settings == null ? null : "Yes";
            sheet.Cells[row, 12].Value2 = settings?.Target;
            sheet.Cells[row, 13].Value2 = settings?.Direction?.ToString();
            sheet.Cells[row, 14].Value2 = settings?.RequestedConfidence;
        }

        private static ForecastTargetSettings? ReadTargetSettings(dynamic sheet, int row, ValidationResult? validation)
        {
            object? marker = ReadOptional(sheet, row, 11, "TargetConfigured");
            object? target = ReadOptional(sheet, row, 12, "Target");
            object? direction = ReadOptional(sheet, row, 13, "TargetDirection");
            object? confidence = ReadOptional(sheet, row, 14, "RequestedConfidence");
            if (marker == null && target == null && direction == null && confidence == null) return null;
            bool valid = string.Equals(Convert.ToString(marker), "Yes", StringComparison.OrdinalIgnoreCase);
            double? targetValue = null, confidenceValue = null;
            TargetDirection? directionValue = null;
            if (target != null)
            {
                if (SimulationValidation.TryNumber(target, out double value)) targetValue = value;
                else valid = false;
            }
            if (confidence != null)
            {
                if (SimulationValidation.TryNumber(confidence, out double value) && value >= 0 && value <= 100 && targetValue.HasValue)
                    confidenceValue = value;
                else valid = false;
            }
            if (!string.IsNullOrEmpty(Convert.ToString(direction)))
            {
                if (Enum.TryParse(Convert.ToString(direction), out TargetDirection value) && Enum.IsDefined(value)) directionValue = value;
                else valid = false;
            }
            if (!valid)
                validation?.Add($"Saved forecast row {row}", "The saved target settings are incomplete or invalid.",
                    "Redefine this forecast in Model Manager and set its target again in Results.");
            return valid ? new ForecastTargetSettings(targetValue, directionValue, confidenceValue) : null;
        }

        private const string LinkPrefix = "_MC_Cell_";

        private static object? ReadOptional(dynamic sheet, int row, int column, string header)
        {
            return string.Equals(Convert.ToString(sheet.Cells[1, column].Value2), header, StringComparison.OrdinalIgnoreCase)
                ? ExcelSimulationWorkbook.ReadCellValue(ExcelDnaUtil.Application, sheet.Cells[row, column]) : null;
        }

        private static string EnsureCellLink(dynamic workbook, string link, string sheet, string address)
        {
            // Never rebind a broken saved name to whatever now occupies its old address.
            if (!string.IsNullOrEmpty(link)) return link;
            try
            {
                dynamic cell = workbook.Worksheets[sheet].Range[address];
                if (Convert.ToDouble(cell.Cells.CountLarge) != 1)
                    throw new InvalidOperationException("A model item must refer to one cell.");
                string name = LinkPrefix + Guid.NewGuid().ToString("N");
                string reference = "='" + ((string)cell.Worksheet.Name).Replace("'", "''") + "'!" + cell.Address[true, true];
                workbook.Names.Add(Name: name, RefersTo: reference, Visible: false);
                return name;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Older/stale rows remain intact and will be reported on restore/simulation.
                return "";
            }
        }

        private static void ResolveSavedLink(dynamic workbook, string link, ref string sheet, ref string address,
            ValidationResult? validation, string location)
        {
            try
            {
                dynamic cell;
                if (string.IsNullOrEmpty(link)) cell = workbook.Worksheets[sheet].Range[address];
                else
                {
                    if (!link.StartsWith(LinkPrefix, StringComparison.Ordinal))
                        throw new InvalidOperationException("Invalid saved link.");
                    cell = workbook.Names[link].RefersToRange;
                }
                if (Convert.ToDouble(cell.Cells.CountLarge) != 1 ||
                    !string.Equals((string)cell.Worksheet.Parent.Name, (string)workbook.Name, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Invalid saved range.");
                sheet = cell.Worksheet.Name;
                address = cell.Address[false, false];
            }
            catch (Exception ex) when (ex is System.Runtime.InteropServices.COMException || ex is InvalidOperationException)
            {
                validation?.Add(location, "The saved worksheet or cell no longer exists or cannot be resolved.",
                    "Select the intended cell and redefine this item, or remove it in Model Manager.");
                // A broken tracked reference must not fall back to a plausible but different cell.
                if (!string.IsNullOrEmpty(link)) address = "#REF!";
            }
        }

        private static dynamic GetOrCreateConfigSheet(
            dynamic workbook)
        {
            dynamic existingSheet =
                FindConfigSheet(
                    workbook);


            if (existingSheet != null)
            {
                return
                    existingSheet;
            }


            dynamic newSheet =
                workbook.Worksheets.Add();


            newSheet.Name =
                ConfigSheetName;


            return
                newSheet;
        }


        // =========================================================
        // FIND CONFIG SHEET
        // =========================================================

        private static dynamic FindConfigSheet(
            dynamic workbook)
        {
            int sheetCount =
                workbook.Worksheets.Count;


            for (
                int i = 1;
                i <= sheetCount;
                i++)
            {
                dynamic sheet =
                    workbook.Worksheets[
                        i];


                string sheetName =
                    Convert.ToString(
                        sheet.Name)
                    ?? "";


                if (
                    string.Equals(
                        sheetName,
                        ConfigSheetName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return
                        sheet;
                }
            }


            return
                null;
        }


        // =========================================================
        // SAFE DOUBLE CONVERSION
        // =========================================================

        private static double GetDoubleValue(
            object value)
        {
            if (value == null)
            {
                return
                    0;
            }


            try
            {
                return
                    Convert.ToDouble(
                        value);
            }
            catch
            {
                return
                    0;
            }
        }
    }
}
