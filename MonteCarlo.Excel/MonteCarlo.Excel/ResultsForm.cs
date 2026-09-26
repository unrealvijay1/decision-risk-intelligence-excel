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

        private ForecastRunResult currentForecastResult;

        private readonly ComboBox cmbForecast;

        private readonly Label lblTargetSummary = new Label();
        private readonly Label lblProbabilityCaption = new Label();
        private readonly Label lblP50 = new Label();
        private readonly Label lblP80 = new Label();
        private readonly Label[] detailValues = new Label[7];
        private readonly ToolTip fullText = new ToolTip();
        private readonly Font primaryFont = new Font("Segoe UI", 28);
        private readonly Font neutralFont = new Font("Segoe UI", 16);

        private readonly Panel histogramPanel;
        private readonly Panel tornadoPanel;

        private static bool cumulativeViewSelected;
        private readonly ComboBox cmbChartView;
        private readonly System.Collections.Generic.Dictionary<ForecastRunResult, CumulativeProbabilityResult>
            cumulativeResults = new();

        private readonly TextBox txtTarget;
        private readonly TextBox txtConfidence;
        private readonly Label lblRequiredTarget;
        private enum TargetOrigin { Manual, Confidence }
        private TargetOrigin targetOrigin;
        private double? acceptedConfidence;
        private bool updatingTargetText;
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

            cmbChartView = new ComboBox
            {
                Left = 850, Top = 120, Width = 210,
                DropDownStyle = ComboBoxStyle.DropDownList,
                AccessibleName = "Forecast chart view"
            };
            cmbChartView.Items.AddRange(new object[] { "Distribution", "Cumulative" });
            cmbChartView.SelectedIndex = cumulativeViewSelected ? 1 : 0;
            lblHistogram.Text = cumulativeViewSelected ? "Cumulative Probability" : "Forecast Distribution";
            cmbChartView.SelectedIndexChanged += (_, _) =>
            {
                cumulativeViewSelected = cmbChartView.SelectedIndex == 1;
                lblHistogram.Text = cumulativeViewSelected ? "Cumulative Probability" : "Forecast Distribution";
                histogramPanel.Invalidate();
            };
            Controls.Add(cmbChartView);


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
                if (updatingTargetText) return;
                targetOrigin = TargetOrigin.Manual;
                acceptedConfidence = null;
                if (lblRequiredTarget != null) lblRequiredTarget.Text = "";
                currentTarget = null;
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
                RefreshSuccessDisplay();
                histogramPanel.Invalidate();
            };

            var calculationTabs = new TabControl
            {
                Left = 30, Top = 515, Width = 435, Height = 126
            };
            var targetTab = new TabPage("Target → Confidence");
            var confidenceTab = new TabPage("Confidence → Target");
            calculationTabs.TabPages.AddRange(new[] { targetTab, confidenceTab });
            Controls.Add(calculationTabs);
            targetTab.Controls.AddRange(new Control[]
                { lblTarget, txtTarget, btnCalculate });
            lblTarget.SetBounds(5, 9, 95, 24);
            txtTarget.SetBounds(105, 5, 180, 24);
            btnCalculate.SetBounds(300, 3, 100, 28);

            confidenceTab.Controls.Add(new Label
                { Text = "Confidence Level", Left = 5, Top = 9, Width = 115, Height = 24 });
            txtConfidence = new TextBox
                { Text = "80", Left = 125, Top = 5, Width = 115, AccessibleName = "Confidence Level (%)" };
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
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, Padding = new Padding(24, 12, 24, 12),
                ColumnCount = 2, RowCount = 6
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 57));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 43));
            foreach (int height in new[] { 40, 38, 112 })
                root.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 62));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 38));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            void Place(TableLayoutPanel parent, Control control, int column, int row)
            {
                control.Dock = DockStyle.Fill;
                control.Margin = new Padding(3);
                parent.Controls.Add(control, column, row);
            }
            Place(root, title, 0, 0);
            root.SetColumnSpan(title, 2);
            var forecast = Grid(2, 1);
            forecast.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85));
            forecast.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Place(forecast, forecastLabel, 0, 0);
            Place(forecast, cmbForecast, 1, 0);
            cmbForecast.DropDownWidth = 900;
            Place(root, forecast, 0, 1);
            root.SetColumnSpan(forecast, 2);

            var summary = Grid(4, 1);
            foreach (int share in new[] { 38, 26, 18, 18 })
                summary.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, share));
            var primary = Grid(1, 2);
            primary.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
            primary.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            Place(primary, lblSuccess, 0, 0);
            Place(primary, lblProbabilityCaption, 0, 1);
            Place(summary, primary, 0, 0);
            Label[] landmarks = { lblTargetSummary, lblP50, lblP80 };
            string[] captions = { "Target", "P50", "P80" };
            for (int i = 0; i < landmarks.Length; i++)
            {
                var block = Grid(1, 2);
                block.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
                block.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                Place(block, new Label { Text = captions[i] }, 0, 0);
                landmarks[i].Font = neutralFont;
                landmarks[i].AutoEllipsis = true;
                Place(block, landmarks[i], 0, 1);
                Place(summary, block, i + 1, 0);
            }
            Place(root, summary, 0, 2);
            root.SetColumnSpan(summary, 2);

            var chart = Grid(1, 3);
            chart.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            chart.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            chart.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            var chartHeader = Grid(2, 1);
            chartHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            chartHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 175));
            Place(chartHeader, chartTitle, 0, 0);
            Place(chartHeader, cmbChartView, 1, 0);
            Place(chart, chartHeader, 0, 0);
            Place(chart, histogramPanel, 0, 1);
            var legends = Grid(2, 1);
            legends.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            legends.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            Place(legends, lblSuccessLegend, 0, 0);
            Place(legends, lblMissLegend, 1, 0);
            Place(chart, legends, 0, 2);
            Place(root, chart, 0, 3);
            histogramPanel.Resize += (_, _) => histogramPanel.Invalidate();

            var decision = Grid(1, 5);
            decision.Padding = new Padding(12, 0, 0, 0);
            foreach (int height in new[] { 28, 22, 34 })
                decision.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
            decision.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            decision.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            Place(decision, new Label { Text = "Decision tools", Font = new Font("Segoe UI", 11) }, 0, 0);
            Place(decision, new Label { Text = "Success when" }, 0, 1);
            Place(decision, cmbSuccessDirection, 0, 2);
            Place(decision, calculationTabs, 0, 3);
            // Keep this footer space empty to preserve the approved lower-section layout.
            foreach (TabPage page in calculationTabs.TabPages)
            {
                var input = page.Controls.OfType<TextBox>().Single();
                var button = page.Controls.OfType<Button>().Single();
                var inputLabel = page.Controls.OfType<Label>().First();
                var body = Grid(2, 3);
                body.Padding = new Padding(6);
                body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
                body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
                body.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
                body.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
                body.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                inputLabel.Text = input == txtConfidence ? "Confidence Level (%)" : "Target Value";
                Place(body, inputLabel, 0, 0);
                body.SetColumnSpan(inputLabel, 2);
                Place(body, input, 0, 1);
                Place(body, button, 1, 1);
                Label output = input == txtConfidence ? lblRequiredTarget : new Label
                    { Text = "Actual probability appears in the summary above.", AutoEllipsis = true };
                Place(body, output, 0, 2);
                body.SetColumnSpan(output, 2);
                foreach (Control unused in page.Controls.Cast<Control>().ToArray()) unused.Dispose();
                page.Controls.Add(body);
            }
            Place(root, decision, 1, 3);

            var sensitivity = Grid(1, 2);
            sensitivity.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            sensitivity.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            sensitivityTitle.Font = new Font("Segoe UI", 9);
            Place(sensitivity, sensitivityTitle, 0, 0);
            Place(sensitivity, tornadoPanel, 0, 1);
            tornadoPanel.Resize += (_, _) => tornadoPanel.Invalidate();
            Place(root, sensitivity, 0, 4);
            var details = Grid(2, 8);
            details.Padding = new Padding(12, 0, 0, 0);
            details.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            details.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58));
            for (int i = 0; i < 8; i++) details.RowStyles.Add(new RowStyle(SizeType.Percent, 12.5f));
            Place(details, new Label { Text = "Details" }, 0, 0);
            string[] names = { "Trials", "Mean", "Std Dev", "Minimum", "Maximum", "P10", "P90" };
            for (int i = 0; i < names.Length; i++)
            {
                Place(details, new Label { Text = names[i] }, 0, i + 1);
                detailValues[i] = new Label { AutoEllipsis = true, TextAlign = ContentAlignment.TopRight };
                Place(details, detailValues[i], 1, i + 1);
            }
            Place(root, details, 1, 4);
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            buttons.Controls.AddRange(new Control[] { export, close });
            Place(root, buttons, 0, 5);
            root.SetColumnSpan(buttons, 2);
            foreach (Control unused in oldControls)
                if (unused.Parent == this) unused.Dispose();
            Controls.Add(root);
            ResumeLayout(true);
        }

        private static TableLayoutPanel Grid(int columns, int rows) => new TableLayoutPanel
            { Dock = DockStyle.Fill, ColumnCount = columns, RowCount = rows, Margin = new Padding(0) };

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
            targetOrigin = TargetOrigin.Manual;
            acceptedConfidence = null;
            lblRequiredTarget.Text = "";
            currentTarget = null;
            currentDirection = forecastDirections.TryGetValue(currentForecastResult, out var direction)
                ? direction : null;
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
                FormatValue(currentForecastResult.P10), FormatValue(currentForecastResult.P90) };
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


            CalculateProbability(
                null,
                EventArgs.Empty);


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

            // Keep probability semantics unchanged, but never plot non-finite targets.
            targetOrigin = TargetOrigin.Manual;
            acceptedConfidence = null;
            lblRequiredTarget.Text = "";
            currentTarget = double.IsFinite(target) ? target : null;
            RefreshSuccessDisplay();

        }

        private void CalculateConfidenceTarget()
        {
            if (!double.TryParse(txtConfidence.Text, out double confidence) ||
                !double.IsFinite(confidence) || confidence < 0 || confidence > 100)
            {
                MessageBox.Show("Enter a confidence level from 0 to 100.", "Monte Carlo",
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
            try { txtTarget.Text = target.ToString("R", CultureInfo.CurrentCulture); }
            finally { updatingTargetText = false; }
            currentTarget = target;
            targetOrigin = TargetOrigin.Confidence;
            acceptedConfidence = confidence;
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
            if (cmbChartView.SelectedIndex == 1)
            {
                DrawCumulative(e.Graphics, histogramPanel.ClientRectangle);
                return;
            }
            DrawHistogram(
                e.Graphics,
                histogramPanel.ClientRectangle);
        }

        private void DrawCumulative(Graphics graphics, Rectangle area)
        {
            var state = graphics.Save();
            try
            {
                graphics.SetClip(area);
                graphics.Clear(SystemColors.Window);
                if (!cumulativeResults.TryGetValue(currentForecastResult, out var data))
                {
                    data = CumulativeProbability.Build(currentForecastResult.Values);
                    cumulativeResults.Add(currentForecastResult, data);
                }
                using Font axisFont = new Font("Segoe UI", 8);
                if (data.ValidationMessage != null || data.Points.Count == 0)
                {
                    graphics.DrawString(data.ValidationMessage ?? "No simulation results available.",
                        axisFont, SystemBrushes.ControlText, new RectangleF(12, 12, area.Width - 24, area.Height - 24));
                    return;
                }

                double min = data.Points[0].X;
                double max = data.Points[^1].X;
                if (currentTarget.HasValue)
                {
                    min = Math.Min(min, currentTarget.Value);
                    max = Math.Max(max, currentTarget.Value);
                }
                // Padding makes the 0%/100% tails and constant-outcome jump visible.
                double padding = Math.Max(Math.Abs(min), Math.Abs(max)) * .05;
                if (padding == 0) padding = 1;
                double paddedMin = min - padding;
                double paddedMax = max + padding;
                if (double.IsFinite(paddedMin)) min = paddedMin;
                if (double.IsFinite(paddedMax)) max = paddedMax;

                using Font markerFont = new Font("Segoe UI", 8, FontStyle.Bold);
                int rowHeight = (int)Math.Ceiling(markerFont.GetHeight(graphics)) + 8;
                int top = 4 + 3 * rowHeight + 4;
                const int left = 55;
                int width = area.Width - left - 25;
                int height = area.Height - top - 55;
                if (width <= 0 || height <= 0) return;
                float bottom = top + height;
                float Y(double probability) => bottom - (float)(probability * height);

                DrawTargetRegions(graphics, min, max, left, top, width, height);
                using Pen gridPen = new Pen(SystemColors.ControlLight);
                using Pen axisPen = new Pen(SystemColors.ControlText);
                for (int percent = 0; percent <= 100; percent += 25)
                {
                    float y = Y(percent / 100d);
                    graphics.DrawLine(gridPen, left, y, left + width, y);
                    string label = $"{percent}%";
                    SizeF size = graphics.MeasureString(label, axisFont);
                    graphics.DrawString(label, axisFont, SystemBrushes.ControlText,
                        left - size.Width - 5, y - size.Height / 2);
                }
                graphics.DrawLine(axisPen, left, top, left, bottom);
                graphics.DrawLine(axisPen, left, bottom, left + width, bottom);
                using Pen curvePen = new Pen(SystemColors.Highlight, 2);
                float previousX = left;
                foreach (var point in data.Points)
                {
                    float x = ValueToX(point.X, min, max, left, width);
                    graphics.DrawLine(curvePen, previousX, Y(point.ProbabilityBefore), x, Y(point.ProbabilityBefore));
                    graphics.DrawLine(curvePen, x, Y(point.ProbabilityBefore), x, Y(point.Probability));
                    previousX = x;
                }
                graphics.DrawLine(curvePen, previousX, top, left + width, top);

                // Identical stored percentiles and marker geometry in both views.
                DrawHistogramMarkers(graphics, area, min, max, left, top, width, height,
                    markerFont, 4, rowHeight);
                if (currentTarget.HasValue)
                {
                    double lowerTail = currentForecastResult.ProbabilityLessThanOrEqual(currentTarget.Value);
                    float targetX = ValueToX(currentTarget.Value, min, max, left, width);
                    float guideY = Y(lowerTail);
                    using Pen guidePen = new Pen(Color.DarkOrange, 1.5f) { DashStyle = DashStyle.Dash };
                    graphics.DrawLine(guidePen, left, guideY, targetX, guideY);
                    using Brush dotBrush = new SolidBrush(Color.DarkOrange);
                    graphics.FillEllipse(dotBrush, targetX - 3, guideY - 3, 6, 6);
                    string guideLabel = $"≤ target: {lowerTail:P1}";
                    SizeF labelSize = graphics.MeasureString(guideLabel, axisFont);
                    float labelY = guideY - labelSize.Height - 3;
                    if (labelY < top) labelY = guideY + 3;
                    float labelX = Math.Clamp(targetX - labelSize.Width - 6, left + 3,
                        Math.Max(left + 3, left + width - labelSize.Width));
                    graphics.FillRectangle(SystemBrushes.Window, labelX, labelY, labelSize.Width, labelSize.Height);
                    graphics.DrawString(guideLabel, axisFont, SystemBrushes.ControlText, labelX, labelY);
                }
                DrawXAxisLabel(graphics, min, min, max, left, width, (int)bottom + 8);
                DrawXAxisLabel(graphics, max, min, max, left, width, (int)bottom + 8);
            }
            finally
            {
                graphics.Restore(state);
            }
        }

        // Shared presentation refresh for all target, direction and forecast transitions.
        private void RefreshSuccessDisplay()
        {
            lblSuccessLegend.Visible = false;
            lblMissLegend.Visible = false;
            lblRequiredTarget.Text = "";
            lblSuccess.Font = neutralFont;
            lblSuccess.Text = "No target defined";
            lblProbabilityCaption.Text = "Enter a target or choose a confidence level.";
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
                    lblProbabilityCaption.Text = $"Probability of achieving {comparison} {FormatValue(target)}";
                    lblSuccessLegend.Text = $"Success: {comparison} target";
                    lblMissLegend.Text = currentDirection == TargetDirection.AtOrBelow ? "Miss: > target" : "Miss: < target";
                    lblSuccessLegend.Visible = true;
                    lblMissLegend.Visible = true;
                }
                if (targetOrigin == TargetOrigin.Confidence && acceptedConfidence.HasValue)
                {
                    string description = currentDirection.HasValue ? "confidence" : "cumulative probability";
                    lblRequiredTarget.Text = $"Requested {description}: {acceptedConfidence.Value:0.########}%\n" +
                        $"Target at {acceptedConfidence.Value:0.########}% {description}: {FormatValue(target)}";
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
                fullText.Dispose();
                primaryFont.Dispose();
                neutralFont.Dispose();
            }
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
                graphics.DrawLine(targetPen, targetX.Value - 4, top, targetX.Value + 4, top);
            }

            // All lines precede text. Separate rows identify even exactly coincident markers.
            DrawMarkerLabel(graphics, area, font,
                "P50", p50X,
                headerTop, rowHeight, left, SystemColors.ControlDarkDark);
            DrawMarkerLabel(graphics, area, font,
                "P80", p80X,
                headerTop + rowHeight, rowHeight, left, SystemColors.ControlDarkDark);

            if (targetPosition.Range == ChartValueRange.Invalid) return;
            string targetLabel = "Target";
            if (targetPosition.Range == ChartValueRange.BelowRange)
                targetLabel += " — below simulated range";
            else if (targetPosition.Range == ChartValueRange.AboveRange)
                targetLabel += " — above simulated range";
            DrawMarkerLabel(graphics, area, font, targetLabel, targetX,
                headerTop + 2 * rowHeight, rowHeight, left, Color.DarkOrange);
        }

        private static void DrawMarkerLabel(Graphics graphics, Rectangle area, Font font,
            string text, float? markerX, int rowTop, int rowHeight, int left, Color markerColor)
        {
            using StringFormat format = new StringFormat
            {
                Trimming = StringTrimming.EllipsisCharacter,
                FormatFlags = StringFormatFlags.NoWrap
            };
            float availableWidth = Math.Max(0, area.Width - 4);
            float labelWidth = Math.Min(graphics.MeasureString(text, font).Width + 4, availableWidth);
            float labelX = Math.Clamp(markerX.HasValue ? markerX.Value + 5 : left,
                area.Left + 2, area.Right - 2 - labelWidth);
            using Brush textBrush = new SolidBrush(SystemColors.ControlText);
            graphics.DrawString(text, font, textBrush,
                new RectangleF(labelX, rowTop, labelWidth, rowHeight - 5), format);

            if (!markerX.HasValue) return;
            // Tick remains at the exact X coordinate, below its label even at panel edges.
            float tickY = rowTop + rowHeight - 3;
            using Pen tickPen = new Pen(markerColor, 1);
            float labelAnchor = Math.Clamp(markerX.Value, labelX, labelX + labelWidth);
            graphics.DrawLine(tickPen, labelAnchor, tickY, markerX.Value, tickY);
            graphics.DrawLine(tickPen, markerX.Value, tickY - 2, markerX.Value, tickY + 2);
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
