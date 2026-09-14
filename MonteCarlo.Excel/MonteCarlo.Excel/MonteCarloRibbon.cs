using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace MonteCarlo.Excel
{
    [ComVisible(true)]
    public class MonteCarloRibbon : ExcelRibbon
    {
        // =========================================================
        // RIBBON UI
        // =========================================================

        public override string GetCustomUI(
            string ribbonID)
        {
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
  <ribbon>
    <tabs>

      <tab id='MonteCarloTab'
           label='Monte Carlo'>

        <group id='ModelSetupGroup'
               label='Model Setup'>

          <button
              id='DefineAssumptionButton'
              label='Define Assumption'
              screentip='Define or edit an assumption'
              supertip='Define an uncertain input, manually select a distribution, or fit a distribution from historical data.'
              size='large'
              getImage='GetRibbonImage'
              onAction='OnDefineAssumption'/>

          <button
              id='DefineForecastButton'
              label='Define Forecast'
              screentip='Define or edit a forecast'
              supertip='Define the selected formula cell as a Monte Carlo forecast.'
              size='large'
              getImage='GetRibbonImage'
              onAction='OnDefineForecast'/>

          <button
              id='ModelManagerButton'
              label='Model Manager'
              screentip='Manage Monte Carlo model'
              supertip='View, edit, delete, and navigate to assumptions and forecasts.'
              size='large'
              getImage='GetRibbonImage'
              onAction='OnModelManager'/>

        </group>


        <group id='SimulationGroup'
               label='Simulation'>

          <button
              id='RunSimulationButton'
              label='Run Simulation'
              screentip='Run Monte Carlo simulation'
              supertip='Run the simulation for all assumptions and forecasts and open the results dashboard.'
              size='large'
              getImage='GetRibbonImage'
              onAction='OnRunSimulation'/>

          <button
              id='ReloadModelButton'
              label='Reload Model'
              screentip='Reload saved model'
              supertip='Reload Monte Carlo assumptions and forecasts stored in the current workbook.'
              size='normal'
              getImage='GetRibbonImage'
              onAction='OnReloadModel'/>

        </group>


        <group id='MaintenanceGroup'
               label='Maintenance'>

          <button
              id='ClearModelButton'
              label='Clear Model'
              screentip='Clear Monte Carlo model'
              supertip='Delete all saved Monte Carlo assumptions and forecasts from the current workbook.'
              size='normal'
              getImage='GetRibbonImage'
              onAction='OnClearModel'/>

        </group>

      </tab>

    </tabs>
  </ribbon>
</customUI>";
        }


        // =========================================================
        // CUSTOM RIBBON ICONS
        // =========================================================

        public Bitmap? GetRibbonImage(
            IRibbonControl control)
        {
            string resourceName =
                control.Id switch
                {
                    "DefineAssumptionButton" =>
                        "MonteCarlo.Excel.Icons.assumption.png",

                    "DefineForecastButton" =>
                        "MonteCarlo.Excel.Icons.forecast.png",

                    "ModelManagerButton" =>
                        "MonteCarlo.Excel.Icons.model_manager.png",

                    "RunSimulationButton" =>
                        "MonteCarlo.Excel.Icons.run_simulation.png",

                    "ReloadModelButton" =>
                        "MonteCarlo.Excel.Icons.reload_model.png",

                    "ClearModelButton" =>
                        "MonteCarlo.Excel.Icons.clear_model.png",

                    _ =>
                        ""
                };


            if (string.IsNullOrWhiteSpace(
                    resourceName))
            {
                return null;
            }


            Assembly assembly =
                Assembly.GetExecutingAssembly();


            using Stream? stream =
                assembly.GetManifestResourceStream(
                    resourceName);


            if (stream == null)
            {
                return null;
            }


            using System.Drawing.Image image =
                System.Drawing.Image.FromStream(
                    stream);


            return
                new Bitmap(
                    image);
        }


        // =========================================================
        // DEFINE / EDIT ASSUMPTION
        // =========================================================

        public void OnDefineAssumption(
            IRibbonControl control)
        {
            try
            {
                WorkbookPersistence.LoadModel();


                dynamic excelApp =
                    ExcelDnaUtil.Application;


                dynamic worksheet =
                    excelApp.ActiveSheet;


                dynamic selectedCell =
                    excelApp.Selection;


                if (selectedCell.Cells.Count > 1)
                {
                    MessageBox.Show(
                        "Please select only one cell.",
                        "Monte Carlo");

                    return;
                }


                string address =
                    selectedCell.Address[
                        false,
                        false];


                string sheetName =
                    worksheet.Name;


                AssumptionDefinition? existingAssumption =
                    SimulationModel.FindAssumption(
                        sheetName,
                        address);


                using AssumptionForm form =
                    new AssumptionForm(
                        $"{sheetName}!{address}",
                        existingAssumption);


                if (form.ShowDialog()
                    != DialogResult.OK)
                {
                    return;
                }


                bool wasExisting =
                    existingAssumption != null;


                SimulationModel.Assumptions.RemoveAll(
                    x =>
                        string.Equals(
                            x.SheetName,
                            sheetName,
                            StringComparison.OrdinalIgnoreCase)
                        &&
                        string.Equals(
                            x.CellAddress,
                            address,
                            StringComparison.OrdinalIgnoreCase));


                SimulationModel.Assumptions.Add(
                    new AssumptionDefinition
                    {
                        Name =
                            form.AssumptionName,

                        SheetName =
                            sheetName,

                        CellAddress =
                            address,

                        Distribution =
                            form.Distribution,

                        Parameter1 =
                            form.Parameter1,

                        Parameter2 =
                            form.Parameter2,

                        Parameter3 =
                            form.Parameter3
                    });


                WorkbookPersistence.SaveModel();


                MessageBox.Show(
                    wasExisting
                        ? $"Assumption updated.\n\n" +
                          $"Name: {form.AssumptionName}\n" +
                          $"Cell: {sheetName}!{address}\n" +
                          $"Distribution: {form.Distribution}"
                        : $"Assumption created.\n\n" +
                          $"Name: {form.AssumptionName}\n" +
                          $"Cell: {sheetName}!{address}\n" +
                          $"Distribution: {form.Distribution}",
                    "Monte Carlo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError(
                    ex);
            }
        }


        // =========================================================
        // DEFINE / EDIT FORECAST
        // =========================================================

        public void OnDefineForecast(
            IRibbonControl control)
        {
            try
            {
                WorkbookPersistence.LoadModel();


                dynamic excelApp =
                    ExcelDnaUtil.Application;


                dynamic worksheet =
                    excelApp.ActiveSheet;


                dynamic selectedCell =
                    excelApp.Selection;


                if (selectedCell.Cells.Count > 1)
                {
                    MessageBox.Show(
                        "Please select only one forecast cell.",
                        "Monte Carlo");

                    return;
                }


                string address =
                    selectedCell.Address[
                        false,
                        false];


                string sheetName =
                    worksheet.Name;


                ForecastDefinition? existingForecast =
                    SimulationModel.FindForecast(
                        sheetName,
                        address);


                string defaultName =
                    existingForecast != null &&
                    !string.IsNullOrWhiteSpace(
                        existingForecast.Name)
                        ? existingForecast.Name
                        : $"{sheetName}!{address}";


                string? forecastName =
                    ShowForecastNameDialog(
                        existingForecast == null
                            ? "Define Forecast"
                            : "Edit Forecast",
                        defaultName);


                if (forecastName == null)
                {
                    return;
                }


                forecastName =
                    forecastName.Trim();


                if (string.IsNullOrWhiteSpace(
                    forecastName))
                {
                    forecastName =
                        $"{sheetName}!{address}";
                }


                bool wasExisting =
                    existingForecast != null;


                SimulationModel.Forecasts.RemoveAll(
                    x =>
                        string.Equals(
                            x.SheetName,
                            sheetName,
                            StringComparison.OrdinalIgnoreCase)
                        &&
                        string.Equals(
                            x.CellAddress,
                            address,
                            StringComparison.OrdinalIgnoreCase));


                SimulationModel.Forecasts.Add(
                    new ForecastDefinition
                    {
                        Name =
                            forecastName,

                        SheetName =
                            sheetName,

                        CellAddress =
                            address
                    });


                WorkbookPersistence.SaveModel();


                MessageBox.Show(
                    wasExisting
                        ? $"Forecast updated.\n\n" +
                          $"Name: {forecastName}\n" +
                          $"Cell: {sheetName}!{address}"
                        : $"Forecast added.\n\n" +
                          $"Name: {forecastName}\n" +
                          $"Cell: {sheetName}!{address}",
                    "Monte Carlo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError(
                    ex);
            }
        }


        // =========================================================
        // MODEL MANAGER
        // =========================================================

        public void OnModelManager(
            IRibbonControl control)
        {
            try
            {
                WorkbookPersistence.LoadModel();


                using ModelManagerForm form =
                    new ModelManagerForm();


                form.ShowDialog();
            }
            catch (Exception ex)
            {
                ShowError(
                    ex);
            }
        }


        // =========================================================
        // RELOAD MODEL
        // =========================================================

        public void OnReloadModel(
            IRibbonControl control)
        {
            try
            {
                WorkbookPersistence.LoadModel();


                string assumptionText =
                    SimulationModel.Assumptions.Count == 0
                        ? "None"
                        : string.Join(
                            "\n",
                            SimulationModel.Assumptions.Select(
                                assumption =>
                                    $"{GetAssumptionDisplayName(assumption)} " +
                                    $"({assumption.SheetName}!" +
                                    $"{assumption.CellAddress})"));


                string forecastText =
                    SimulationModel.Forecasts.Count == 0
                        ? "None"
                        : string.Join(
                            "\n",
                            SimulationModel.Forecasts.Select(
                                forecast =>
                                    $"{GetForecastDisplayName(forecast)} " +
                                    $"({forecast.SheetName}!" +
                                    $"{forecast.CellAddress})"));


                MessageBox.Show(
                    $"Model loaded from workbook.\n\n" +
                    $"Assumptions: {SimulationModel.Assumptions.Count}\n" +
                    $"{assumptionText}\n\n" +
                    $"Forecasts: {SimulationModel.Forecasts.Count}\n" +
                    $"{forecastText}",
                    "Monte Carlo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError(
                    ex);
            }
        }


        // =========================================================
        // RUN SIMULATION
        // =========================================================

        public void OnRunSimulation(
            IRibbonControl control)
        {
            dynamic excelApp =
                null;


            try
            {
                WorkbookPersistence.LoadModel();


                if (SimulationModel.Assumptions.Count == 0)
                {
                    MessageBox.Show(
                        "No saved assumptions were found.",
                        "Monte Carlo");

                    return;
                }


                if (SimulationModel.Forecasts.Count == 0)
                {
                    MessageBox.Show(
                        "No saved forecasts were found.",
                        "Monte Carlo");

                    return;
                }


                excelApp =
                    ExcelDnaUtil.Application;


                using SimulationSettingsForm settingsForm =
                    new SimulationSettingsForm();


                if (settingsForm.ShowDialog()
                    != DialogResult.OK)
                {
                    return;
                }


                int trials =
                    settingsForm.Trials;


                if (trials <= 0)
                {
                    return;
                }


                // =================================================
                // ASSUMPTION SENSITIVITY NAMES
                // =================================================

                var assumptionSensitivityNames =
                    new Dictionary<
                        AssumptionDefinition,
                        string>();


                foreach (
                    AssumptionDefinition assumption
                    in SimulationModel.Assumptions)
                {
                    string desiredName =
                        GetAssumptionDisplayName(
                            assumption);


                    string finalName =
                        desiredName;


                    int suffix =
                        2;


                    while (
                        assumptionSensitivityNames
                            .Values
                            .Any(
                                x =>
                                    string.Equals(
                                        x,
                                        finalName,
                                        StringComparison.OrdinalIgnoreCase)))
                    {
                        finalName =
                            $"{desiredName} ({suffix})";


                        suffix++;
                    }


                    assumptionSensitivityNames[
                        assumption] =
                        finalName;
                }


                // =================================================
                // ASSUMPTION SAMPLES
                // =================================================

                var assumptionSamples =
                    new Dictionary<
                        string,
                        double[]>();


                foreach (
                    AssumptionDefinition assumption
                    in SimulationModel.Assumptions)
                {
                    assumptionSamples[
                        assumptionSensitivityNames[
                            assumption]] =
                        new double[trials];
                }


                // =================================================
                // FORECAST SAMPLES
                // =================================================

                var forecastSamples =
                    new Dictionary<
                        ForecastDefinition,
                        double[]>();


                foreach (
                    ForecastDefinition forecast
                    in SimulationModel.Forecasts)
                {
                    forecastSamples[
                        forecast] =
                        new double[trials];
                }


                // =================================================
                // SAVE ORIGINAL ASSUMPTION VALUES
                // =================================================

                var originalValues =
                    new Dictionary<
                        AssumptionDefinition,
                        object?>();


                foreach (
                    AssumptionDefinition assumption
                    in SimulationModel.Assumptions)
                {
                    dynamic sheet =
                        excelApp.Worksheets[
                            assumption.SheetName];


                    dynamic cell =
                        sheet.Range[
                            assumption.CellAddress];


                    originalValues[
                        assumption] =
                        cell.Value2;
                }


                bool originalScreenUpdating =
                    excelApp.ScreenUpdating;


                bool originalEnableEvents =
                    excelApp.EnableEvents;


                excelApp.ScreenUpdating =
                    false;


                excelApp.EnableEvents =
                    false;


                int progressInterval =
                    Math.Max(
                        1,
                        trials / 100);


                // =================================================
                // SIMULATION LOOP
                // =================================================

                try
                {
                    for (
                        int i = 0;
                        i < trials;
                        i++)
                    {
                        // -----------------------------------------
                        // SAMPLE ALL ASSUMPTIONS
                        // -----------------------------------------

                        foreach (
                            AssumptionDefinition assumption
                            in SimulationModel.Assumptions)
                        {
                            dynamic sheet =
                                excelApp.Worksheets[
                                    assumption.SheetName];


                            dynamic cell =
                                sheet.Range[
                                    assumption.CellAddress];


                            double sample =
                                GenerateSample(
                                    assumption);


                            string sensitivityName =
                                assumptionSensitivityNames[
                                    assumption];


                            assumptionSamples[
                                sensitivityName][i] =
                                sample;


                            cell.Value2 =
                                sample;
                        }


                        // -----------------------------------------
                        // CALCULATE MODEL
                        // -----------------------------------------

                        excelApp.Calculate();


                        // -----------------------------------------
                        // CAPTURE ALL FORECASTS
                        // -----------------------------------------

                        foreach (
                            ForecastDefinition forecast
                            in SimulationModel.Forecasts)
                        {
                            dynamic forecastSheet =
                                excelApp.Worksheets[
                                    forecast.SheetName];


                            dynamic forecastCell =
                                forecastSheet.Range[
                                    forecast.CellAddress];


                            object forecastValue =
                                forecastCell.Value2;


                            if (forecastValue == null)
                            {
                                throw new Exception(
                                    $"Forecast " +
                                    $"{forecast.SheetName}!" +
                                    $"{forecast.CellAddress} " +
                                    $"returned no value.");
                            }


                            forecastSamples[
                                forecast][i] =
                                Convert.ToDouble(
                                    forecastValue);
                        }


                        // -----------------------------------------
                        // PROGRESS
                        // -----------------------------------------

                        if (
                            i % progressInterval == 0
                            ||
                            i == trials - 1)
                        {
                            int percentage =
                                (int)(
                                    (i + 1) *
                                    100.0 /
                                    trials);


                            excelApp.StatusBar =
                                $"Monte Carlo Simulation: " +
                                $"{percentage}% " +
                                $"({i + 1:N0}/{trials:N0})";
                        }
                    }
                }
                finally
                {
                    // ---------------------------------------------
                    // RESTORE INPUTS
                    // ---------------------------------------------

                    foreach (
                        var item
                        in originalValues)
                    {
                        dynamic sheet =
                            excelApp.Worksheets[
                                item.Key.SheetName];


                        dynamic cell =
                            sheet.Range[
                                item.Key.CellAddress];


                        cell.Value2 =
                            item.Value;
                    }


                    excelApp.Calculate();


                    excelApp.ScreenUpdating =
                        originalScreenUpdating;


                    excelApp.EnableEvents =
                        originalEnableEvents;


                    excelApp.StatusBar =
                        false;
                }


                // =================================================
                // RESULT OBJECT
                // =================================================

                SimulationRunResult simulationResult =
                    new SimulationRunResult(
                        trials,
                        forecastSamples,
                        assumptionSamples);


                // =================================================
                // RESULTS DASHBOARD
                // =================================================

                using ResultsForm resultsForm =
                    new ResultsForm(
                        simulationResult);


                resultsForm.ShowDialog();
            }
            catch (Exception ex)
            {
                ShowError(
                    ex);
            }
            finally
            {
                if (excelApp != null)
                {
                    try
                    {
                        excelApp.StatusBar =
                            false;
                    }
                    catch
                    {
                    }
                }
            }
        }


        // =========================================================
        // CLEAR MODEL
        // =========================================================

        public void OnClearModel(
            IRibbonControl control)
        {
            try
            {
                DialogResult confirmation =
                    MessageBox.Show(
                        "Clear all Monte Carlo assumptions and " +
                        "forecast definitions from this workbook?",
                        "Monte Carlo",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);


                if (confirmation !=
                    DialogResult.Yes)
                {
                    return;
                }


                SimulationModel.Clear();


                WorkbookPersistence.ClearSavedModel();


                MessageBox.Show(
                    "Monte Carlo model cleared.",
                    "Monte Carlo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError(
                    ex);
            }
        }


        // =========================================================
        // ASSUMPTION DISPLAY NAME
        // =========================================================

        private static string GetAssumptionDisplayName(
            AssumptionDefinition assumption)
        {
            return
                string.IsNullOrWhiteSpace(
                    assumption.Name)
                    ? $"{assumption.SheetName}!" +
                      $"{assumption.CellAddress}"
                    : assumption.Name;
        }


        // =========================================================
        // FORECAST DISPLAY NAME
        // =========================================================

        private static string GetForecastDisplayName(
            ForecastDefinition forecast)
        {
            return
                string.IsNullOrWhiteSpace(
                    forecast.Name)
                    ? $"{forecast.SheetName}!" +
                      $"{forecast.CellAddress}"
                    : forecast.Name;
        }


        // =========================================================
        // FORECAST NAME DIALOG
        // =========================================================

        private static string? ShowForecastNameDialog(
            string title,
            string currentValue)
        {
            using Form dialog =
                new Form
                {
                    Text =
                        title,

                    Width =
                        420,

                    Height =
                        190,

                    StartPosition =
                        FormStartPosition.CenterScreen,

                    FormBorderStyle =
                        FormBorderStyle.FixedDialog,

                    MaximizeBox =
                        false,

                    MinimizeBox =
                        false
                };


            Label label =
                new Label
                {
                    Text =
                        "Forecast name:",

                    Left =
                        20,

                    Top =
                        30,

                    Width =
                        110
                };


            TextBox textBox =
                new TextBox
                {
                    Left =
                        140,

                    Top =
                        25,

                    Width =
                        230,

                    Text =
                        currentValue
                };


            Button btnOK =
                new Button
                {
                    Text =
                        "OK",

                    Left =
                        205,

                    Top =
                        80,

                    Width =
                        75,

                    DialogResult =
                        DialogResult.OK
                };


            Button btnCancel =
                new Button
                {
                    Text =
                        "Cancel",

                    Left =
                        295,

                    Top =
                        80,

                    Width =
                        75,

                    DialogResult =
                        DialogResult.Cancel
                };


            dialog.Controls.Add(
                label);

            dialog.Controls.Add(
                textBox);

            dialog.Controls.Add(
                btnOK);

            dialog.Controls.Add(
                btnCancel);


            dialog.AcceptButton =
                btnOK;


            dialog.CancelButton =
                btnCancel;


            if (dialog.ShowDialog()
                != DialogResult.OK)
            {
                return
                    null;
            }


            return
                textBox.Text;
        }


        // =========================================================
        // DISTRIBUTION SAMPLING
        // =========================================================

        private static double GenerateSample(
            AssumptionDefinition assumption)
        {
            switch (assumption.Distribution)
            {
                case DistributionType.Normal:

                    return
                        MonteCarlo.Core
                            .NormalDistribution
                            .Sample(
                                assumption.Parameter1,
                                assumption.Parameter2);


                case DistributionType.Triangular:

                    return
                        MonteCarlo.Core
                            .TriangularDistribution
                            .Sample(
                                assumption.Parameter1,
                                assumption.Parameter2,
                                assumption.Parameter3);


                case DistributionType.Pert:

                    return
                        MonteCarlo.Core
                            .PertDistribution
                            .Sample(
                                assumption.Parameter1,
                                assumption.Parameter2,
                                assumption.Parameter3);


                case DistributionType.Uniform:

                    return
                        MonteCarlo.Core
                            .UniformDistribution
                            .Sample(
                                assumption.Parameter1,
                                assumption.Parameter2);


                case DistributionType.Lognormal:

                    return
                        MonteCarlo.Core
                            .LognormalDistribution
                            .Sample(
                                assumption.Parameter1,
                                assumption.Parameter2);


                default:

                    throw new InvalidOperationException(
                        $"Unsupported distribution: " +
                        $"{assumption.Distribution}");
            }
        }


        // =========================================================
        // ERROR HANDLER
        // =========================================================

        private static void ShowError(
            Exception ex)
        {
            MessageBox.Show(
                ex.ToString(),
                "Monte Carlo Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}