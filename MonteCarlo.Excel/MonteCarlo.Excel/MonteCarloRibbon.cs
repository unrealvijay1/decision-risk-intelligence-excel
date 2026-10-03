using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

using MonteCarlo.Excel.Licensing;

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

          <button id='CorrelationsButton' label='Correlations' size='large' getImage='GetRibbonImage' onAction='OnCorrelations'/>

          <button
              id='ModelManagerButton'
              label='Model Manager'
              screentip='Manage Monte Carlo model'
              supertip='View, edit, delete, and navigate to assumptions and forecasts.'
              size='large'
              getImage='GetRibbonImage'
              onAction='OnModelManager'/>

          <button
              id='ClearCellDefinitionButton'
              label='Clear Cell Definition'
              screentip='Remove the selected cell definition'
              supertip='Remove the Monte Carlo definition from one selected cell and restore its original fill. The cell value or formula is preserved.'
              size='large'
              getImage='GetRibbonImage'
              onAction='OnClearCellDefinition'/>

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
          <button id='ScenarioAnalysisButton' label='Scenario Analysis' size='large' getImage='GetRibbonImage' onAction='OnScenarioAnalysis'/>
          <button id='OptimizerButton' label='Optimizer' size='large' getImage='GetRibbonImage' onAction='OnOptimizer'/>
          <button id='SimulationSettingsButton' label='Simulation Settings' size='large' getImage='GetRibbonImage' onAction='OnSimulationSettings'/>

          <button
              id='ReloadModelButton'
              label='Reload Model'
              screentip='Reload saved model'
              supertip='Reload Monte Carlo assumptions and forecasts stored in the current workbook.'
              size='large'
              getImage='GetRibbonImage'
              onAction='OnReloadModel'/>

        </group>


        <group id='ProcessAnalysisGroup' label='Process Analysis'>
          <button id='SpcAnalysisButton' label='SPC Analysis' size='large' getImage='GetRibbonImage'
                  screentip='Analyze process behaviour with XmR charts' onAction='OnSpcAnalysis'/>
        </group>

        <group id='MaintenanceGroup'
               label='Maintenance'>

          <button
              id='ClearModelButton'
              label='Clear Model'
              screentip='Clear Monte Carlo model'
              supertip='Delete all saved Monte Carlo assumptions and forecasts from the current workbook.'
              size='large'
              getImage='GetRibbonImage'
              onAction='OnClearModel'/>

        </group>


        <group id='ProductGroup'
               label='Product'>

          <button
              id='LicenseButton'
              label='License'
              screentip='View license information'
              supertip='View your Monte Carlo license type, activation status and expiry information.'
              size='large'
              imageMso='FileProperties'
              onAction='OnLicense'/>

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

                    "CorrelationsButton" => "MonteCarlo.Excel.Icons.Correlations.png",
                    "ClearCellDefinitionButton" => "MonteCarlo.Excel.Icons.Clear Cell Defition.png",
                    "ScenarioAnalysisButton" => "MonteCarlo.Excel.Icons.Scenario Analysis.png",
                    "OptimizerButton" => "MonteCarlo.Excel.Icons.Optimizer.png",
                    "SimulationSettingsButton" => "MonteCarlo.Excel.Icons.Simulation Settings.png",
                    "SpcAnalysisButton" => "MonteCarlo.Excel.Icons.SPC Analysis.png",

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


            // Fit supplied artwork into a native large-control canvas, without upscaling.
            var bitmap = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using var graphics = Graphics.FromImage(bitmap);
            graphics.Clear(Color.Transparent);
            double scale = Math.Min(1d, Math.Min(32d / image.Width, 32d / image.Height));
            int width = Math.Max(1, (int)Math.Round(image.Width * scale));
            int height = Math.Max(1, (int)Math.Round(image.Height * scale));
            if (scale == 1d)
                graphics.DrawImageUnscaled(image, (32 - width) / 2, (32 - height) / 2);
            else
            {
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                using var attributes = new System.Drawing.Imaging.ImageAttributes();
                attributes.SetWrapMode(System.Drawing.Drawing2D.WrapMode.TileFlipXY);
                graphics.DrawImage(image, new Rectangle((32 - width) / 2, (32 - height) / 2, width, height),
                    0, 0, image.Width, image.Height, GraphicsUnit.Pixel, attributes);
            }
            return bitmap;
        }


        // =========================================================
        // DEFINE / EDIT ASSUMPTION
        // =========================================================

        public void OnSimulationSettings(IRibbonControl control)
        {
            try
            {
                var validation = new ValidationResult();
                var settings = WorkbookPersistence.LoadSimulationSettings(validation);
                if (!validation.IsValid) MessageBox.Show(validation.UserMessage, "Simulation Settings");
                using var form = new SimulationSettingsForm(settings);
                if (form.ShowDialog() == DialogResult.OK) WorkbookPersistence.SaveSimulationSettings(form.Settings);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Simulation Settings"); }
        }
        public void OnClearCellDefinition(IRibbonControl control)
        {
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                var result = ModelDefinitionDeletion.ClearSelectedCell((object?)app.Selection,
                    (title, message) => MessageBox.Show(message, title, MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes);
                if (result == DeleteDefinitionResult.InvalidSelection)
                    MessageBox.Show(ModelDefinitionDeletion.SelectionMessage, "Monte Carlo");
                else if (result == DeleteDefinitionResult.NoDefinition)
                    MessageBox.Show(ModelDefinitionDeletion.NoDefinitionMessage, "Monte Carlo");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Monte Carlo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
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
                        CellLink = existingAssumption?.CellLink ?? "",
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
                            form.Parameter3,

                        Parameter4 =
                            form.Parameter4,
                        ProbabilityTable = form.ProbabilityTable
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
                    existingForecast != null
                    &&
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
                        CellLink = existingForecast?.CellLink ?? "",
                        TargetSettings = existingForecast?.TargetSettings,
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

        public void OnCorrelations(IRibbonControl control)
        {
            try
            {
                LicenseService.EnsureAccess();
                dynamic app = ExcelDnaUtil.Application;
                object workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("Open your model workbook first.");
                WorkbookPersistence.LoadModelForSimulation(restoreHighlights: false);
                var assumptions = SimulationModel.Assumptions.ToArray();
                using var form = new AssumptionCorrelationsForm(assumptions, WorkbookPersistence.LoadCorrelations(workbook),
                    values => WorkbookPersistence.SaveCorrelations(workbook, assumptions, values));
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError(ex.ToString());
                MessageBox.Show(ex.Message, "Assumption Correlations", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void OnScenarioAnalysis(IRibbonControl control)
        {
            try { ScenarioAnalysisCommand.Show(); }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError(ex.ToString());
                MessageBox.Show(ex.Message, "Scenario Analysis", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void OnOptimizer(IRibbonControl control)
        {
            try { OptimizerCommand.Show(); }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError(ex.ToString());
                MessageBox.Show(ex.Message, "Optimizer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void OnSpcAnalysis(IRibbonControl control)
        {
            try { SpcAnalysisCommand.Show(); }
            catch (ArgumentException ex) { MessageBox.Show(ex.Message, "SPC Analysis", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError(ex.ToString());
                MessageBox.Show("SPC Analysis could not open. Check your license and that the workbook is available.", "SPC Analysis", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

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
        // LICENSE
        // =========================================================

        public void OnLicense(
            IRibbonControl control)
        {
            try
            {
                using LicenseForm form =
                    new LicenseForm();


                form.ShowDialog();
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
            try
            {
                // =================================================
                // LICENSE CHECK
                // =================================================

                LicenseInfo license =
                    LicenseService.GetCurrentLicense();


                if (!license.IsValid)
                {
                    MessageBox.Show(
                        "Your Monte Carlo license is not active.\n\n" +
                        "Please activate or renew your license " +
                        "to run simulations.",
                        "Monte Carlo License",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);


                    using LicenseForm licenseForm =
                        new LicenseForm();


                    licenseForm.ShowDialog();


                    return;
                }


                // =================================================
                // SIMULATION SETTINGS
                // =================================================

                using var cancellation = new CancellationTokenSource();
                using var live = new LiveRunForm("Monte Carlo — live simulation", cancellation);
                dynamic liveApp = ExcelDnaUtil.Application;
                bool interactive = liveApp.Interactive;
                SimulationExecutionResult outcome;
                live.Show();
                try
                {
                    liveApp.Interactive = false;
                    outcome = SimulationService.TryRunLive(live.Trial, cancellation.Token);
                }
                finally { live.Finish(); liveApp.Interactive = interactive; }
                if (outcome.DiagnosticException is OperationCanceledException && outcome.RestorationFailures.Count == 0)
                    return;
                if (!outcome.Succeeded)
                {
                    if (outcome.DiagnosticException != null)
                        System.Diagnostics.Trace.TraceError(outcome.DiagnosticException.ToString());
                    MessageBox.Show(outcome.UserMessage, "Monte Carlo — Check simulation",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                SimulationRunResult simulationResult = outcome.Result!;


                // =================================================
                // RESULTS
                // =================================================

                using ResultsForm resultsForm =
                    new ResultsForm(
                        simulationResult);


                resultsForm.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError(ex.ToString());
                MessageBox.Show("The simulation could not finish. Check that the workbook is open and available, then try again.",
                    "Monte Carlo", MessageBoxButtons.OK, MessageBoxIcon.Error);
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


            textBox.SelectAll();


            if (dialog.ShowDialog()
                != DialogResult.OK)
            {
                return null;
            }


            return
                textBox.Text;
        }


        // =========================================================
        // ERROR
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
