using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using MonteCarlo.Core;

namespace MonteCarlo.Excel
{
    public class ResultsForm : Form
    {
        private readonly SimulationRunResult simulationResult;
        private readonly object targetWorkbook;

        private ForecastRunResult currentForecastResult;

        private readonly ComboBox cmbForecast;

        private readonly Label lblTargetSummary = new Label();
        private readonly Label lblProbabilityCaption = new Label();
        private readonly Label lblP50 = new Label();
        private readonly Label lblP80 = new Label();
        private readonly Label[] detailValues = new Label[9];
        private readonly ToolTip fullText = new ToolTip();
        private readonly Font primaryFont = new Font("Segoe UI", 28);
        private readonly Font neutralFont = new Font("Segoe UI", 16);

        private readonly Panel histogramPanel;
        private readonly Panel tornadoPanel;



        private readonly TextBox txtTarget;
        private readonly TextBox txtConfidence;
        private readonly Label lblRequiredTarget;
        private readonly Label lblCalculatedProbability = new Label();
        private double? displayedTargetValue;
        private enum TargetOrigin { Manual, Confidence }
        private TargetOrigin targetOrigin;
        private double? acceptedConfidence;
        private bool updatingTargetText;
        private bool restoringTargetSettings;
        private double? currentTarget;
        private TargetDirection? currentDirection;
        private readonly System.Collections.Generic.Dictionary<ForecastRunResult, TargetDirection>
            forecastDirections = new();
        private readonly ComboBox cmbSuccessDirection;
        private readonly Label lblSuccess;
        private readonly Label lblSuccessLegend;
        private readonly Label lblMissLegend;
        private bool restoringDirection;
        private static readonly Color SuccessRegionColor = Color.FromArgb(226, 242, 230);
        private static readonly Color MissRegionColor = Color.FromArgb(250, 231, 234);



        public ResultsForm(
            SimulationRunResult result)
        {
            dynamic application = ExcelDna.Integration.ExcelDnaUtil.Application;
            targetWorkbook = application.ActiveWorkbook;
            simulationResult =
                result;


            if (simulationResult.ForecastResults.Count == 0)
            {
                throw new ArgumentException(
                    "Simulation contains no forecast results.");
            }


            currentForecastResult =
                simulationResult.ForecastResults[0];


            // =====================================================
            // FORM
            // =====================================================

            Text =
                "Monte Carlo Simulation Results";

            Width =
                1120;

            Height =
                830;

            StartPosition =
                FormStartPosition.CenterScreen;

            FormBorderStyle =
                FormBorderStyle.FixedDialog;

            MaximizeBox =
                false;

            MinimizeBox =
                false;


            // =====================================================
            // TITLE
            // =====================================================

            Label lblTitle =
                new Label
                {
                    Text =
                        "Monte Carlo Simulation Results",

                    Left =
                        30,

                    Top =
                        20,

                    Width =
                        500,

                    Height =
                        40,

                    Font =
                        new Font(
                            "Segoe UI",
                            18,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblTitle);


            // =====================================================
            // FORECAST SELECTOR
            // =====================================================

            Label lblForecast =
                new Label
                {
                    Text =
                        "Forecast",

                    Left =
                        30,

                    Top =
                        75,

                    Width =
                        80,

                    Height =
                        25,

                    Font =
                        new Font(
                            "Segoe UI",
                            10,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblForecast);


            cmbForecast =
                new ComboBox
                {
                    Left =
                        115,

                    Top =
                        70,

                    Width =
                        320,

                    DropDownStyle =
                        ComboBoxStyle.DropDownList
                };


            foreach (
                ForecastRunResult forecastResult
                in simulationResult.ForecastResults)
            {
                string displayName =
                    GetForecastDisplayName(
                        forecastResult);


                cmbForecast.Items.Add(
                    displayName);
            }


            cmbForecast.SelectedIndexChanged +=
                ForecastSelectionChanged;


            Controls.Add(
                cmbForecast);


            // =====================================================
            // HISTOGRAM HEADER
            // =====================================================

            Label lblHistogram =
                new Label
                {
                    Text =
                        "Forecast Distribution",

                    Left =
                        500,

                    Top =
                        125,

                    Width =
                        280,

                    Height =
                        25,

                    Font =
                        new Font(
                            "Segoe UI",
                            11,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblHistogram);


            // =====================================================
            // HISTOGRAM PANEL
            // =====================================================

            histogramPanel =
                new Panel
                {
                    Left =
                        500,

                    Top =
                        160,

                    Width =
                        560,

                    Height =
                        285,

                    BorderStyle =
                        BorderStyle.FixedSingle
                };


            histogramPanel.Paint +=
                HistogramPanel_Paint;


            Controls.Add(
                histogramPanel);




            // =====================================================
            // PROBABILITY ANALYSIS
            // =====================================================

            Label lblProbabilityHeader =
                new Label
                {
                    Text =
                        "Probability Analysis",

                    Left =
                        30,

                    Top =
                        485,

                    Width =
                        230,

                    Height =
                        25,

                    Font =
                        new Font(
                            "Segoe UI",
                            11,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblProbabilityHeader);


            Label lblTarget =
                new Label
                {
                    Text =
                        "Target Value",

                    Left =
                        30,

                    Top =
                        530,

                    Width =
                        100,

                    Height =
                        25
                };


            Controls.Add(
                lblTarget);


            txtTarget =
                new TextBox
                {
                    Left =
                        140,

                    Top =
                        525,

                    Width =
                        180
                };


            Controls.Add(
                txtTarget);

            // An edited value is not the accepted probability target until Calculate.
            txtTarget.TextChanged += (_, _) =>
            {
                if (updatingTargetText || restoringTargetSettings) return;
                displayedTargetValue = null;
                targetOrigin = TargetOrigin.Manual;
                acceptedConfidence = null;
                if (lblRequiredTarget != null) lblRequiredTarget.Text = "";
                currentTarget = null;
                if (string.IsNullOrWhiteSpace(txtTarget.Text)) PersistTargetSettings();
                RefreshSuccessDisplay();
                histogramPanel.Invalidate();
            };


            Button btnCalculate =
                new Button
                {
                    Text =
                        "Calculate",

                    Left =
                        335,

                    Top =
                        523,

                    Width =
                        100,

                    Height =
                        30
                };


            btnCalculate.Click +=
                CalculateProbability;


            Controls.Add(
                btnCalculate);


            Controls.Add(new Label
            {
                Text = "Success when:", Left = 30, Top = 648, Width = 110, Height = 24
            });
            cmbSuccessDirection = new ComboBox
            {
                Left = 140, Top = 645, Width = 295,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbSuccessDirection.Items.AddRange(new object[]
            {
                "At or below target (≤)", "At or above target (≥)"
            });
            Controls.Add(cmbSuccessDirection);
            lblSuccess = new Label
            {
                Left = 30, Top = 677, Width = 430, Height = 25,
                Font = new Font("Segoe UI", 10, FontStyle.Bold), AutoEllipsis = true
            };
            lblSuccessLegend = new Label
            {
                Left = 30, Top = 705, Width = 205, Height = 22,
                BackColor = SuccessRegionColor, Visible = false
            };
            lblMissLegend = new Label
            {
                Left = 245, Top = 705, Width = 190, Height = 22,
                BackColor = MissRegionColor, Visible = false
            };
            Controls.Add(lblSuccess);
            Controls.Add(lblSuccessLegend);
            Controls.Add(lblMissLegend);
            cmbSuccessDirection.SelectedIndexChanged += (_, _) =>
            {
                if (restoringDirection) return;
                TargetDirection? previousDirection = currentDirection;
                currentDirection = cmbSuccessDirection.SelectedIndex switch
                {
                    0 => TargetDirection.AtOrBelow,
                    1 => TargetDirection.AtOrAbove,
                    _ => null
                };
                if (targetOrigin == TargetOrigin.Confidence && acceptedConfidence.HasValue &&
                    !ApplyConfidenceTarget(acceptedConfidence.Value))
                {
                    currentDirection = previousDirection;
                    restoringDirection = true;
                    cmbSuccessDirection.SelectedIndex = previousDirection switch
                    {
                        TargetDirection.AtOrBelow => 0, TargetDirection.AtOrAbove => 1, _ => -1
                    };
                    restoringDirection = false;
                    return;
                }
                if (currentDirection.HasValue)
                    forecastDirections[currentForecastResult] = currentDirection.Value;
                else
                    forecastDirections.Remove(currentForecastResult);
                PersistTargetSettings();
                RefreshSuccessDisplay();
                histogramPanel.Invalidate();
            };

            var calculationTabs = new TabControl
            {
                Left = 30, Top = 515, Width = 435, Height = 126
            };
            var targetTab = new TabPage("Target → Probability");
            var confidenceTab = new TabPage("Probability → Target");
            calculationTabs.TabPages.AddRange(new[] { targetTab, confidenceTab });
            Controls.Add(calculationTabs);
            targetTab.Controls.AddRange(new Control[]
                { lblTarget, txtTarget, btnCalculate });
            lblTarget.SetBounds(5, 9, 95, 24);
            txtTarget.SetBounds(105, 5, 180, 24);
            btnCalculate.SetBounds(300, 3, 100, 28);

            confidenceTab.Controls.Add(new Label
                { Text = "Probability", Left = 5, Top = 9, Width = 115, Height = 24 });
            txtConfidence = new TextBox
                { Text = "80", Left = 125, Top = 5, Width = 115, AccessibleName = "Probability (%)" };
            confidenceTab.Controls.Add(txtConfidence);
            confidenceTab.Controls.Add(new Label
                { Text = "%", Left = 245, Top = 9, Width = 25, Height = 24 });
            var btnConfidence = new Button
                { Text = "Calculate", Left = 300, Top = 3, Width = 100, Height = 28 };
            btnConfidence.Click += (_, _) => CalculateConfidenceTarget();
            confidenceTab.Controls.Add(btnConfidence);
            lblRequiredTarget = new Label
                { Left = 5, Top = 35, Width = 410, Height = 64, AutoEllipsis = true };
            confidenceTab.Controls.Add(lblRequiredTarget);


            // =====================================================
            // SENSITIVITY ANALYSIS
            // =====================================================

            Label lblSensitivity =
                new Label
                {
                    Text =
                        "Sensitivity Analysis",

                    Left =
                        500,

                    Top =
                        485,

                    Width =
                        280,

                    Height =
                        25,

                    Font =
                        new Font(
                            "Segoe UI",
                            11,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblSensitivity);


            // =====================================================
            // TORNADO PANEL
            // =====================================================

            tornadoPanel =
                new Panel
                {
                    Left =
                        500,

                    Top =
                        520,

                    Width =
                        560,

                    Height =
                        200,

                    BorderStyle =
                        BorderStyle.FixedSingle
                };


            tornadoPanel.Paint +=
                TornadoPanel_Paint;


            Controls.Add(
                tornadoPanel);


            // =====================================================
            // EXPORT REPORT BUTTON
            // =====================================================

            Button btnExportReport =
                new Button
                {
                    Text =
                        "Export Report",

                    Left =
                        30,

                    Top =
                        735,

                    Width =
                        140,

                    Height =
                        34
                };


            btnExportReport.Click +=
                (_, _) =>
                {
                    try
                    {
                        ReportExporter.Export(
                            simulationResult);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            ex.ToString(),
                            "Monte Carlo Export Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                };


            Controls.Add(
                btnExportReport);


            // =====================================================
            // CLOSE BUTTON
            // =====================================================

            Button btnClose =
                new Button
                {
                    Text =
                        "Close",

                    Left =
                        185,

                    Top =
                        735,

                    Width =
                        100,

                    Height =
                        34
                };


            btnClose.Click +=
                (_, _) =>
                {
                    Close();
                };


            Controls.Add(
                btnClose);


            // =====================================================
            // INITIAL FORECAST
            // =====================================================

            BuildDecisionLayout(lblTitle, lblForecast, lblHistogram, calculationTabs,
                lblSensitivity, btnExportReport, btnClose);

            cmbForecast.SelectedIndex =
                0;


            RefreshForecastDisplay();
        }


        // =========================================================
        // FORECAST SELECTION CHANGED
        // =========================================================

        private void BuildDecisionLayout(Label title, Label forecastLabel, Label chartTitle,
            TabControl calculationTabs, Label sensitivityTitle, Button export, Button close)
        {
            SuspendLayout();
            var oldControls = Controls.Cast<Control>().ToArray();
            void Place(TableLayoutPanel parent, Control control, int column, int row)
            {
                control.Dock = DockStyle.Fill;
                control.Margin = new Padding(3);
                parent.Controls.Add(control, column, row);
            }
            var root = Grid(1, 4);
            root.Padding = new Padding(20, 12, 20, 12);
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            title.AutoSize = true;
            Place(root, title, 0, 0);
            var forecast = Grid(2, 1);
            forecast.AutoSize = true;
            forecast.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            forecast.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            forecastLabel.AutoSize = true;
            forecastLabel.TextAlign = ContentAlignment.MiddleLeft;
            Place(forecast, forecastLabel, 0, 0);
            Place(forecast, cmbForecast, 1, 0);
            cmbForecast.DropDownWidth = 900;
            Place(root, forecast, 0, 1);

            var tabs = new TabControl { Name = "ResultsTabs", Dock = DockStyle.Fill };
            var forecastPage = new TabPage("Forecast");
            var sensitivityPage = new TabPage("Sensitivity");
            var statisticsPage = new TabPage("Statistics");
            tabs.TabPages.AddRange(new[] { forecastPage, sensitivityPage, statisticsPage });
            tabs.SelectedIndex = 0;
            Place(root, tabs, 0, 2);

            var forecastBody = Grid(2, 1);
            forecastBody.Padding = new Padding(8);
            forecastBody.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            forecastBody.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66));
            forecastBody.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var controlsColumn = Grid(1, 3);
            controlsColumn.Name = "ForecastControls";
            controlsColumn.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            controlsColumn.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            controlsColumn.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Place(forecastBody, controlsColumn, 0, 0);
            forecastPage.Controls.Add(forecastBody);
            var summary = Grid(1, 2);
            summary.AutoSize = true;
            summary.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            summary.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var probability = Grid(1, 2);
            probability.AutoSize = true;
            probability.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            probability.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            lblSuccess.AutoSize = lblProbabilityCaption.AutoSize = true;
            Place(probability, lblSuccess, 0, 0);
            Place(probability, lblProbabilityCaption, 0, 1);
            Place(summary, probability, 0, 0);
            var target = Grid(1, 2);
            target.AutoSize = true;
            target.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            target.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Place(target, new Label { Text = "Target", AutoSize = true }, 0, 0);
            lblTargetSummary.Font = neutralFont;
            lblTargetSummary.AutoSize = true;
            Place(target, lblTargetSummary, 0, 1);
            Place(summary, target, 0, 1);
            Place(controlsColumn, summary, 0, 0);

            var chart = Grid(1, 2);
            chart.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            chart.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var chartHeader = Grid(1, 1);
            chartHeader.AutoSize = true;
            chartHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            chartTitle.AutoSize = true;
            Place(chartHeader, chartTitle, 0, 0);

            Place(chart, chartHeader, 0, 0);
            Place(chart, histogramPanel, 0, 1);
            Place(forecastBody, chart, 1, 0);
            histogramPanel.Resize += (_, _) => histogramPanel.Invalidate();

            var decision = Grid(1, 3);
            decision.AutoSize = true;

            decision.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            decision.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            decision.RowStyles.Add(new RowStyle(SizeType.Absolute, 140));
            var toolsTitle = new Label { Text = "Decision Tools", AutoSize = true };
            Place(decision, toolsTitle, 0, 0);

            var direction = Grid(1, 2);
            direction.AutoSize = true;
            direction.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            direction.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            Place(direction, new Label { Text = "Success when:", AutoSize = true }, 0, 0);
            Place(direction, cmbSuccessDirection, 0, 1);
            Place(decision, direction, 0, 1);
            foreach (TabPage page in calculationTabs.TabPages)
            {
                var input = page.Controls.OfType<TextBox>().Single();
                var button = page.Controls.OfType<Button>().Single();
                var inputLabel = page.Controls.OfType<Label>().First();
                var body = Grid(2, 3);
                body.AutoSize = true;
                body.Padding = new Padding(4);
                page.BackColor = SystemColors.Control;
                body.BackColor = SystemColors.Control;
                body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
                body.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                for (int i = 0; i < 3; i++) body.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                inputLabel.Text = input == txtConfidence ? "Probability (%)" : "Target Value";
                inputLabel.AutoSize = true;
                button.AutoSize = true;
                button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                Place(body, inputLabel, 0, 0);
                body.SetColumnSpan(inputLabel, 2);
                Place(body, input, 0, 1);
                Place(body, button, 1, 1);
                Label output = input == txtConfidence ? lblRequiredTarget : lblCalculatedProbability;
                output.AutoSize = true;
                Place(body, output, 0, 2);
                body.SetColumnSpan(output, 2);
                foreach (Control unused in page.Controls.Cast<Control>().ToArray()) unused.Dispose();
                page.Controls.Add(body);
            }
            Place(decision, calculationTabs, 0, 2);
            // TabControl does not auto-size to its pages. Measure their content rather than
            // feeding its previously stretched/scaled bounds back into an AutoSize row.
            decision.Layout += (_, _) =>
            {
                int height = calculationTabs.TabPages.Cast<TabPage>()
                    .Max(page => page.Controls[0].GetPreferredSize(new Size(calculationTabs.Width, 0)).Height)
                    + calculationTabs.ItemSize.Height + 8;
                if (decision.RowStyles[2].Height != height) decision.RowStyles[2].Height = height;
            };
            Place(controlsColumn, decision, 0, 1);

            var sensitivity = Grid(1, 3);
            sensitivity.Padding = new Padding(12);
            sensitivity.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            sensitivity.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            sensitivity.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            sensitivityTitle.AutoSize = true;
            Place(sensitivity, sensitivityTitle, 0, 0);
            Place(sensitivity, new Label { Text = "Shows which assumptions have the greatest influence on the selected forecast.", AutoSize = true }, 0, 1);
            Place(sensitivity, tornadoPanel, 0, 2);
            sensitivityPage.Controls.Add(sensitivity);
            tornadoPanel.Resize += (_, _) => tornadoPanel.Invalidate();

            var statistics = Grid(2, 1);
            statistics.Padding = new Padding(16);
            statistics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));
            statistics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            statisticsPage.Controls.Add(statistics);
            TableLayoutPanel Section(string heading, string[] names, Label[] values)
            {
                var table = Grid(2, names.Length + 1);
                table.Dock = DockStyle.Top;
                table.AutoSize = true;
                table.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                table.Padding = new Padding(8);
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48));
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 52));
                for (int i = 0; i <= names.Length; i++) table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                var label = new Label { Text = heading, AutoSize = true };
                Place(table, label, 0, 0); table.SetColumnSpan(label, 2);
                for (int i = 0; i < names.Length; i++)
                {
                    Place(table, new Label { Text = names[i], AutoSize = true, TextAlign = ContentAlignment.MiddleLeft }, 0, i + 1);
                    values[i].AutoSize = true;
                    values[i].Font = Font;
                    values[i].TextAlign = ContentAlignment.MiddleRight;
                    Place(table, values[i], 1, i + 1);
                }
                return table;
            }
            for (int i = 0; i < detailValues.Length; i++) detailValues[i] = new Label();
            statistics.Controls.Add(Section("Distribution Statistics",
                new[] { "Mean", "Std Dev", "Minimum", "Maximum", "P10", "P50", "P80", "P90" },
                new[] { detailValues[1], detailValues[2], detailValues[3], detailValues[4], detailValues[5], lblP50, lblP80, detailValues[6] }), 0, 0);
            statistics.Controls.Add(Section("Simulation Run", new[] { "Trials", "Seed Mode", "Seed" },
                new[] { detailValues[0], detailValues[8], detailValues[7] }), 1, 0);

            var footer = Grid(2, 1);
            footer.AutoSize = true;
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            export.AutoSize = close.AutoSize = true;
            export.Anchor = AnchorStyles.Left;
            close.Anchor = AnchorStyles.Right;
            footer.Controls.Add(export, 0, 0); footer.Controls.Add(close, 1, 0);
            Place(root, footer, 0, 3);
            tabs.SelectedIndexChanged += (_, _) => { histogramPanel.Invalidate(); tornadoPanel.Invalidate(); };
            foreach (Control unused in oldControls)
                if (unused.Parent == this) { if (unused == lblSuccessLegend || unused == lblMissLegend) Controls.Remove(unused); else unused.Dispose(); }
            Controls.Add(root);
            ResumeLayout(true);
        }

        private static TableLayoutPanel Grid(int columns, int rows)
        {
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = columns, RowCount = rows,
                Margin = new Padding(0), AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            if (columns == 1) grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            return grid;
        }
        private void ForecastSelectionChanged(
            object? sender,
            EventArgs e)
        {
            if (
                cmbForecast.SelectedIndex < 0
                ||
                cmbForecast.SelectedIndex >=
                    simulationResult.ForecastResults.Count)
            {
                return;
            }


            currentForecastResult =
                simulationResult
                    .ForecastResults[
                        cmbForecast.SelectedIndex];


            RefreshForecastDisplay();
        }


        // =========================================================
        // REFRESH DISPLAY
        // =========================================================

        private void RefreshForecastDisplay()
        {
            displayedTargetValue = null;
            restoringTargetSettings = true;
            var savedTarget = currentForecastResult.Forecast.TargetSettings;
            targetOrigin = TargetOrigin.Manual;
            acceptedConfidence = null;
            lblRequiredTarget.Text = "";
            currentTarget = null;
            currentDirection = forecastDirections.TryGetValue(currentForecastResult, out var direction)
                ? direction : savedTarget?.Direction;
            restoringDirection = true;
            cmbSuccessDirection.SelectedIndex = currentDirection switch
            {
                TargetDirection.AtOrBelow => 0,
                TargetDirection.AtOrAbove => 1,
                _ => -1
            };
            restoringDirection = false;
            RefreshSuccessDisplay();

            lblP50.Text = FormatValue(currentForecastResult.P50);
            lblP80.Text = FormatValue(currentForecastResult.P80);
            fullText.SetToolTip(lblP50, lblP50.Text);
            fullText.SetToolTip(lblP80, lblP80.Text);
            fullText.SetToolTip(cmbForecast, GetForecastDisplayName(currentForecastResult));
            string[] details = { currentForecastResult.Trials.ToString("N0"),
                FormatValue(currentForecastResult.Mean), FormatValue(currentForecastResult.StandardDeviation),
                FormatValue(currentForecastResult.Minimum), FormatValue(currentForecastResult.Maximum),
                FormatValue(currentForecastResult.P10), FormatValue(currentForecastResult.P90), simulationResult.ActualSeedUsed?.ToString() ?? "Unavailable", simulationResult.Settings?.SeedMode.ToString() ?? "Unavailable" };
            for (int i = 0; i < details.Length; i++)
            {
                detailValues[i].Text = details[i];
                fullText.SetToolTip(detailValues[i], details[i]);
            }

            txtTarget.Text =
                currentForecastResult
                    .P80
                    .ToString(
                        "0.00",
                        CultureInfo.CurrentCulture);


            if (savedTarget == null) CalculateProbability(null, EventArgs.Empty);
            else
            {
                currentTarget = savedTarget.Target;
                acceptedConfidence = savedTarget.RequestedConfidence;
                targetOrigin = acceptedConfidence.HasValue ? TargetOrigin.Confidence : TargetOrigin.Manual;
                SetTargetDisplay(currentTarget);
                if (acceptedConfidence.HasValue) txtConfidence.Text = acceptedConfidence.Value.ToString("R", CultureInfo.CurrentCulture);
                RefreshSuccessDisplay();
            }
            restoringTargetSettings = false;

            histogramPanel.Invalidate();

            tornadoPanel.Invalidate();
        }


        // =========================================================
        // FORECAST DISPLAY NAME
        // =========================================================

        private static string GetForecastDisplayName(
            ForecastRunResult forecastResult)
        {
            if (
                !string.IsNullOrWhiteSpace(
                    forecastResult.Forecast.Name))
            {
                return
                    forecastResult
                        .Forecast
                        .Name;
            }


            return
                $"{forecastResult.Forecast.SheetName}!" +
                $"{forecastResult.Forecast.CellAddress}";
        }


        // =========================================================
        // PROBABILITY ANALYSIS
        // =========================================================

        private void CalculateProbability(
            object? sender,
            EventArgs e)
        {
            double? exactDisplayedTarget = displayedTargetValue;
            currentTarget = null;
            RefreshSuccessDisplay();
            histogramPanel.Invalidate();

            if (!double.TryParse(
                    txtTarget.Text,
                    out double target))
            {
                MessageBox.Show(
                    "Enter a valid target value.",
                    "Monte Carlo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            target = exactDisplayedTarget ?? target;
            // Keep probability semantics unchanged, but never plot non-finite targets.
            targetOrigin = TargetOrigin.Manual;
            acceptedConfidence = null;
            lblRequiredTarget.Text = "";
            currentTarget = double.IsFinite(target) ? target : null;
            SetTargetDisplay(currentTarget);
            PersistTargetSettings();
            RefreshSuccessDisplay();

        }

        private void SetTargetDisplay(double? target)
        {
            bool previous = updatingTargetText;
            updatingTargetText = true;
            try { txtTarget.Text = target.HasValue ? FormatValue(target.Value) : ""; }
            finally { updatingTargetText = previous; }
            displayedTargetValue = target;
        }
        private void PersistTargetSettings()
        {
            if (restoringTargetSettings) return;
            try
            {
                WorkbookPersistence.SaveTargetSettings(currentForecastResult.Forecast,
                    new ForecastTargetSettings(currentTarget, currentDirection, acceptedConfidence), targetWorkbook);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError("Target settings could not be saved: {0}", ex);
                MessageBox.Show("The target settings could not be saved to this workbook. Check that the workbook is editable and the forecast still exists.",
                    "Monte Carlo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CalculateConfidenceTarget()
        {
            if (!double.TryParse(txtConfidence.Text, out double confidence) ||
                !double.IsFinite(confidence) || confidence < 0 || confidence > 100)
            {
                MessageBox.Show("Enter a probability from 0 to 100.", "Monte Carlo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            ApplyConfidenceTarget(confidence);
        }

        private bool ApplyConfidenceTarget(double confidence)
        {
            double target;
            try
            {
                target = currentForecastResult.GetTargetForConfidence(confidence, currentDirection);
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "Monte Carlo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // Never round-trip display text through the probability calculation.
            updatingTargetText = true;
            try { SetTargetDisplay(target); }
            finally { updatingTargetText = false; }
            currentTarget = target;
            targetOrigin = TargetOrigin.Confidence;
            acceptedConfidence = confidence;
            PersistTargetSettings();
            RefreshSuccessDisplay();
            histogramPanel.Invalidate();
            return true;
        }


        // =========================================================
        // HISTOGRAM PAINT
        // =========================================================

        private void HistogramPanel_Paint(
            object? sender,
            PaintEventArgs e)
        {
            DrawHistogram(
                e.Graphics,
                histogramPanel.ClientRectangle);
        }

        // Shared presentation refresh for all target, direction and forecast transitions.
        private void RefreshSuccessDisplay()
        {
            lblSuccessLegend.Visible = false;
            lblMissLegend.Visible = false;
            lblRequiredTarget.Text = "";
            lblCalculatedProbability.Text = "Probability: —";
            lblSuccess.Font = neutralFont;
            lblSuccess.Text = "No target defined";
            lblProbabilityCaption.Text = "Enter a target or choose a probability.";
            lblTargetSummary.Text = "—";
            if (currentTarget.HasValue)
            {
                double target = currentTarget.Value;
                string comparison = currentDirection == TargetDirection.AtOrAbove ? "≥" : "≤";
                lblTargetSummary.Text = (currentDirection.HasValue ? comparison + " " : "") + FormatValue(target);
                double below = currentForecastResult.ProbabilityLessThanOrEqual(target);
                lblSuccess.Text = "Select success direction";
                lblProbabilityCaption.Text = "Choose at or below, or at or above target.";
                if (currentDirection.HasValue)
                {
                    double probability = currentDirection == TargetDirection.AtOrBelow ? below
                        : currentForecastResult.ProbabilityGreaterThanOrEqual(target);
                    lblSuccess.Font = primaryFont;
                    lblSuccess.Text = $"{probability:P1}";
                    lblCalculatedProbability.Text = $"Probability: {probability:P1}";
                    lblProbabilityCaption.Text = "Probability of achieving target";
                    lblSuccessLegend.Text = $"Success: {comparison} target";
                    lblMissLegend.Text = currentDirection == TargetDirection.AtOrBelow ? "Miss: > target" : "Miss: < target";
                    lblSuccessLegend.Visible = true;
                    lblMissLegend.Visible = true;
                }
                if (targetOrigin == TargetOrigin.Confidence && acceptedConfidence.HasValue)
                {
                    lblRequiredTarget.Text = $"Target: {FormatValue(target)}";
                }
            }
            fullText.SetToolTip(lblTargetSummary, lblTargetSummary.Text);
            fullText.SetToolTip(lblProbabilityCaption, lblProbabilityCaption.Text);
            fullText.SetToolTip(lblRequiredTarget, lblRequiredTarget.Text);
            histogramPanel.Invalidate();
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing)
            {
                lblSuccessLegend.Dispose();
                lblMissLegend.Dispose();
                fullText.Dispose();
                primaryFont.Dispose();
                neutralFont.Dispose();
            }
        }

        private void DrawRegionLabels(Graphics graphics, double min, double max, int left, int top, int width, int height)
        {
            if (!currentDirection.HasValue || !currentTarget.HasValue || width <= 0 || height <= 0) return;
            var position = ChartValuePosition.Calculate(currentTarget, min, max);
            if (position.Range == ChartValueRange.Invalid) return;
            float boundary = position.Range == ChartValueRange.BelowRange ? left :
                position.Range == ChartValueRange.AboveRange ? left + width : ValueToX(currentTarget.Value, min, max, left, width);
            bool below = currentDirection == TargetDirection.AtOrBelow;
            using var font = new Font("Segoe UI", 8);
            void LabelRegion(float x, float regionWidth, bool success)
            {
                string text = success ? "Success" : "Miss";
                SizeF size = graphics.MeasureString(text, font);
                if (regionWidth < size.Width + 8 || height < size.Height + 16) return;
                float labelX = x + (regionWidth - size.Width) / 2;
                float labelY = top + height - size.Height - 12;
                using var background = new SolidBrush(success ? SuccessRegionColor : MissRegionColor);
                graphics.FillRectangle(background, labelX - 2, labelY, size.Width + 4, size.Height);
                graphics.DrawString(text, font, SystemBrushes.ControlText, labelX, labelY);
            }
            LabelRegion(left, boundary - left, below);
            LabelRegion(boundary, left + width - boundary, !below);
        }
        private void DrawTargetRegions(Graphics graphics, double min, double max,
            int left, int top, int width, int height)
        {
            if (!currentDirection.HasValue || width <= 0 || height <= 0) return;
            var position = ChartValuePosition.Calculate(currentTarget, min, max);
            if (position.Range == ChartValueRange.Invalid) return;

            bool leftIsSuccess = currentDirection == TargetDirection.AtOrBelow;
            using Brush success = new SolidBrush(SuccessRegionColor);
            using Brush miss = new SolidBrush(MissRegionColor);
            if (position.Range != ChartValueRange.InRange)
            {
                bool allSuccess = position.Range == ChartValueRange.AboveRange
                    ? leftIsSuccess : !leftIsSuccess;
                graphics.FillRectangle(allSuccess ? success : miss, left, top, width, height);
                return;
            }

            // These are value regions, not estimates of probability within a bin.
            float boundary = ValueToX(currentTarget!.Value, min, max, left, width);
            graphics.FillRectangle(leftIsSuccess ? success : miss,
                left, top, boundary - left, height);
            graphics.FillRectangle(leftIsSuccess ? miss : success,
                boundary, top, left + width - boundary, height);
        }


        // =========================================================
        // HISTOGRAM
        // =========================================================

        private void DrawHistogram(
            Graphics graphics,
            Rectangle area)
        {
            graphics.SetClip(
                area);


            graphics.Clear(
                SystemColors.Window);


            double min =
                currentForecastResult.Minimum;


            double max =
                currentForecastResult.Maximum;


            if (max <= min)
            {
                graphics.ResetClip();

                return;
            }


            int binCount =
                Math.Min(
                    30,
                    Math.Max(
                        15,
                        (int)Math.Sqrt(
                            currentForecastResult.Trials)));


            int[] bins =
                new int[binCount];


            double binWidth =
                (max - min) /
                binCount;


            foreach (
                double value
                in currentForecastResult.Values)
            {
                int index =
                    (int)(
                        (value - min) /
                        binWidth);


                index =
                    Math.Max(
                        0,
                        Math.Min(
                            binCount - 1,
                            index));


                bins[index]++;
            }


            int maxFrequency =
                bins.Max();


            if (maxFrequency <= 0)
            {
                graphics.ResetClip();

                return;
            }


            int leftMargin =
                55;

            int rightMargin =
                25;

            using Font markerFont = new Font("Segoe UI", 8, FontStyle.Bold);
            int markerRowHeight = (int)Math.Ceiling(markerFont.GetHeight(graphics)) + 8;
            const int markerHeaderTop = 4;
            int topMargin = markerHeaderTop + 3 * markerRowHeight + 4;

            int bottomMargin =
                55;


            int chartWidth =
                area.Width -
                leftMargin -
                rightMargin;


            int chartHeight =
                area.Height -
                topMargin -
                bottomMargin;


            using Pen histogramAxisPen =
                new Pen(
                    SystemColors.ControlText);

            DrawTargetRegions(graphics, min, max, leftMargin, topMargin, chartWidth, chartHeight);


            graphics.DrawLine(
                histogramAxisPen,
                leftMargin,
                topMargin,
                leftMargin,
                topMargin +
                    chartHeight);


            graphics.DrawLine(
                histogramAxisPen,
                leftMargin,
                topMargin +
                    chartHeight,
                leftMargin +
                    chartWidth,
                topMargin +
                    chartHeight);


            float histogramBarWidth =
                (float)chartWidth /
                binCount;


            using Brush histogramBarBrush =
                new SolidBrush(
                    SystemColors.Highlight);


            for (
                int i = 0;
                i < binCount;
                i++)
            {
                float barHeight =
                    (float)bins[i] /
                    maxFrequency *
                    chartHeight;


                float x =
                    leftMargin +
                    i *
                    histogramBarWidth;


                float y =
                    topMargin +
                    chartHeight -
                    barHeight;


                graphics.FillRectangle(
                    histogramBarBrush,
                    x + 1,
                    y,
                    Math.Max(
                        1,
                        histogramBarWidth - 2),
                    barHeight);
            }


            DrawRegionLabels(graphics, min, max, leftMargin, topMargin, chartWidth, chartHeight);
            DrawHistogramMarkers(graphics, area, min, max, leftMargin,
                topMargin, chartWidth, chartHeight, markerFont, markerHeaderTop, markerRowHeight);

            // Percentile values are shown in the header; retain only range labels here.
            DrawXAxisLabel(graphics, min, min, max, leftMargin, chartWidth,
                topMargin + chartHeight + 8);
            DrawXAxisLabel(graphics, max, min, max, leftMargin, chartWidth,
                topMargin + chartHeight + 8);

            graphics.ResetClip();
        }

        private void DrawHistogramMarkers(Graphics graphics, Rectangle area,
            double min, double max, int left, int top, int width, int height,
            Font font, int headerTop, int rowHeight)
        {
            if (width <= 0 || height <= 0) return;

            float p50X = ValueToX(currentForecastResult.P50, min, max, left, width);
            float p80X = ValueToX(currentForecastResult.P80, min, max, left, width);
            using Pen percentilePen = new Pen(SystemColors.ControlDarkDark, 1.5f)
            {
                DashStyle = DashStyle.Dash
            };
            using Pen targetPen = new Pen(Color.DarkOrange, 2.5f);
            graphics.DrawLine(percentilePen, p50X, top, p50X, top + height);
            graphics.DrawLine(percentilePen, p80X, top, p80X, top + height);

            var targetPosition = ChartValuePosition.Calculate(currentTarget, min, max);
            float? targetX = null;
            if (targetPosition.Range == ChartValueRange.InRange)
            {
                targetX = ValueToX(currentTarget!.Value, min, max, left, width);
                graphics.DrawLine(targetPen, targetX.Value, top, targetX.Value, top + height);

            }

            // All lines precede text. Separate rows identify even exactly coincident markers.
            DrawMarkerLabel(graphics, area, font,
                MarkerText("P50", currentForecastResult.P50), p50X,
                headerTop, rowHeight, left, top, SystemColors.ControlDarkDark);
            DrawMarkerLabel(graphics, area, font,
                MarkerText("P80", currentForecastResult.P80), p80X,
                headerTop + rowHeight, rowHeight, left, top, SystemColors.ControlDarkDark);

            if (targetPosition.Range == ChartValueRange.Invalid) return;
            string targetLabel = MarkerText("Target", currentTarget!.Value, currentDirection);
            if (targetPosition.Range == ChartValueRange.BelowRange)
                targetLabel += " — below simulated range";
            else if (targetPosition.Range == ChartValueRange.AboveRange)
                targetLabel += " — above simulated range";
            DrawMarkerLabel(graphics, area, font, targetLabel, targetX,
                headerTop + 2 * rowHeight, rowHeight, left, top, Color.DarkOrange);
        }

        private static string MarkerText(string name, double value, TargetDirection? direction = null) =>
            name + " " + (direction == TargetDirection.AtOrBelow ? "≤ " : direction == TargetDirection.AtOrAbove ? "≥ " : "") + FormatValue(value);
        private static RectangleF MarkerLabelBounds(Rectangle area, float measuredWidth,
            float? markerX, int rowTop, int rowHeight, int left)
        {
            float width = Math.Min(measuredWidth, Math.Max(0, area.Width - 4));
            float x = Math.Clamp(markerX.HasValue ? markerX.Value - width / 2 : left,
                area.Left + 2, area.Right - 2 - width);
            return new RectangleF(x, rowTop, width, rowHeight - 5);
        }
        private static void DrawMarkerLabel(Graphics graphics, Rectangle area, Font font,
            string text, float? markerX, int rowTop, int rowHeight, int left, int plotTop, Color markerColor)
        {
            using StringFormat format = new StringFormat
            {
                Trimming = StringTrimming.EllipsisCharacter,
                FormatFlags = StringFormatFlags.NoWrap
            };
            RectangleF bounds = MarkerLabelBounds(area, graphics.MeasureString(text, font).Width + 4,
                markerX, rowTop, rowHeight, left);
            // Extend the actual marker vertically into its label band. No detached ticks
            // or horizontal leaders; fixed rows keep close/coincident values distinct.
            if (markerX.HasValue)
            {
                using var connector = new Pen(markerColor, markerColor == Color.DarkOrange ? 2.5f : 1.5f);
                graphics.DrawLine(connector, markerX.Value, bounds.Bottom, markerX.Value, plotTop);
            }
            graphics.FillRectangle(SystemBrushes.Window, bounds);
            using Brush textBrush = new SolidBrush(markerColor);
            graphics.DrawString(text, font, textBrush, bounds, format);
        }

        // =========================================================
        // X-AXIS LABEL
        // =========================================================

        private void DrawXAxisLabel(
            Graphics graphics,
            double value,
            double min,
            double max,
            int leftMargin,
            int chartWidth,
            int y)
        {
            float x =
                ValueToX(
                    value,
                    min,
                    max,
                    leftMargin,
                    chartWidth);


            string text =
                FormatCompactNumber(
                    value);


            using Font axisFont =
                new Font(
                    "Segoe UI",
                    8);


            using Brush axisBrush =
                new SolidBrush(
                    SystemColors.ControlText);


            SizeF size =
                graphics.MeasureString(
                    text,
                    axisFont);


            float drawX =
                x -
                size.Width /
                2;


            drawX =
                Math.Max(
                    2,
                    Math.Min(
                        histogramPanel.ClientRectangle.Width -
                        size.Width -
                        2,
                        drawX));


            graphics.DrawString(
                text,
                axisFont,
                axisBrush,
                drawX,
                y);
        }


        // =========================================================
        // VALUE TO X POSITION
        // =========================================================

        private static float ValueToX(
            double value,
            double min,
            double max,
            int leftMargin,
            int chartWidth)
        {
            if (max <= min)
            {
                return
                    leftMargin;
            }


            ChartValuePosition position = ChartValuePosition.Calculate(value, min, max);
            // Existing markers retain endpoint clamping; targets are classified before drawing.
            double ratio = position.Range == ChartValueRange.AboveRange
                ? 1 : position.NormalizedPosition ?? 0;


            return
                leftMargin +
                (float)(
                    ratio *
                    chartWidth);
        }


        // =========================================================
        // TORNADO PAINT
        // =========================================================

        private void TornadoPanel_Paint(
            object? sender,
            PaintEventArgs e)
        {
            DrawTornado(
                e.Graphics,
                tornadoPanel.ClientRectangle);
        }


        // =========================================================
        // TORNADO
        // =========================================================

        private void DrawTornado(
            Graphics graphics,
            Rectangle area)
        {
            graphics.SetClip(
                area);


            graphics.Clear(
                SystemColors.Window);


            if (
                currentForecastResult.Sensitivities == null
                ||
                currentForecastResult.Sensitivities.Count == 0)
            {
                using Font emptyFont =
                    new Font(
                        "Segoe UI",
                        9);


                graphics.DrawString(
                    "No sensitivity data available.",
                    emptyFont,
                    SystemBrushes.ControlText,
                    15,
                    15);


                graphics.ResetClip();

                return;
            }


            int topMargin =
                15;

            int bottomMargin =
                30;

            int leftMargin =
                140;

            int rightMargin =
                55;


            int chartWidth =
                area.Width -
                leftMargin -
                rightMargin;


            int chartHeight =
                area.Height -
                topMargin -
                bottomMargin;


            float centerX =
                leftMargin +
                chartWidth /
                2f;


            using Pen tornadoCenterPen =
                new Pen(
                    SystemColors.ControlDark);


            graphics.DrawLine(
                tornadoCenterPen,
                centerX,
                topMargin,
                centerX,
                topMargin +
                    chartHeight);


            int sensitivityCount =
                currentForecastResult
                    .Sensitivities
                    .Count;


            float rowHeight =
                (float)chartHeight /
                sensitivityCount;


            using Brush positiveBarBrush =
                new SolidBrush(
                    SystemColors.Highlight);


            using Brush negativeBarBrush =
                new SolidBrush(
                    SystemColors.ControlDarkDark);


            using Brush tornadoTextBrush =
                new SolidBrush(
                    SystemColors.ControlText);


            using Font tornadoFont =
                new Font(
                    "Segoe UI",
                    8);


            for (
                int i = 0;
                i < sensitivityCount;
                i++)
            {
                SensitivityResult sensitivity =
                    currentForecastResult
                        .Sensitivities[i];


                float y =
                    topMargin +
                    i *
                    rowHeight +
                    4;


                float maxBarWidth =
                    chartWidth /
                    2f -
                    8;


                float sensitivityBarWidth =
                    (float)(
                        Math.Abs(
                            sensitivity.Correlation)
                        *
                        maxBarWidth);


                float sensitivityBarHeight =
                    Math.Max(
                        6,
                        rowHeight -
                        9);


                string assumptionName =
                    sensitivity.Name;


                if (assumptionName.Length > 18)
                {
                    assumptionName =
                        assumptionName.Substring(
                            0,
                            18);
                }


                graphics.DrawString(
                    assumptionName,
                    tornadoFont,
                    tornadoTextBrush,
                    5,
                    y);


                if (
                    sensitivity.Correlation >= 0)
                {
                    graphics.FillRectangle(
                        positiveBarBrush,
                        centerX,
                        y,
                        sensitivityBarWidth,
                        sensitivityBarHeight);
                }
                else
                {
                    graphics.FillRectangle(
                        negativeBarBrush,
                        centerX -
                            sensitivityBarWidth,
                        y,
                        sensitivityBarWidth,
                        sensitivityBarHeight);
                }


                string correlationText =
                    sensitivity
                        .Correlation
                        .ToString(
                            "0.00");


                SizeF correlationSize =
                    graphics.MeasureString(
                        correlationText,
                        tornadoFont);


                float correlationX;


                if (
                    sensitivity.Correlation >= 0)
                {
                    correlationX =
                        centerX +
                        sensitivityBarWidth +
                        5;
                }
                else
                {
                    correlationX =
                        centerX -
                        sensitivityBarWidth -
                        correlationSize.Width -
                        5;
                }


                correlationX =
                    Math.Max(
                        leftMargin,
                        Math.Min(
                            area.Width -
                            correlationSize.Width -
                            3,
                            correlationX));


                graphics.DrawString(
                    correlationText,
                    tornadoFont,
                    tornadoTextBrush,
                    correlationX,
                    y);
            }


            DrawTornadoScaleLabel(
                graphics,
                "-1",
                leftMargin,
                area.Height -
                    bottomMargin +
                    7);


            DrawTornadoScaleLabel(
                graphics,
                "0",
                centerX,
                area.Height -
                    bottomMargin +
                    7);


            DrawTornadoScaleLabel(
                graphics,
                "+1",
                leftMargin +
                    chartWidth,
                area.Height -
                    bottomMargin +
                    7);


            graphics.ResetClip();
        }


        // =========================================================
        // TORNADO SCALE LABEL
        // =========================================================

        private static void DrawTornadoScaleLabel(
            Graphics graphics,
            string text,
            float x,
            float y)
        {
            using Font scaleFont =
                new Font(
                    "Segoe UI",
                    8);


            SizeF labelSize =
                graphics.MeasureString(
                    text,
                    scaleFont);


            graphics.DrawString(
                text,
                scaleFont,
                SystemBrushes.ControlText,
                x -
                    labelSize.Width /
                    2,
                y);
        }


        // =========================================================
        // FORMATTING
        // =========================================================

        private static string FormatValue(
            double value)
        {
            return
                value.ToString(
                    "N2",
                    CultureInfo.CurrentCulture);
        }


        private static string FormatCompactNumber(
            double value)
        {
            double absolute =
                Math.Abs(
                    value);


            if (absolute >= 10000000)
            {
                return
                    (value / 10000000.0)
                    .ToString("0.0") +
                    "Cr";
            }


            if (absolute >= 100000)
            {
                return
                    (value / 100000.0)
                    .ToString("0.0") +
                    "L";
            }


            if (absolute >= 1000)
            {
                return
                    (value / 1000.0)
                    .ToString("0.0") +
                    "K";
            }


            return
                value.ToString(
                    "0");
        }
    }
}
