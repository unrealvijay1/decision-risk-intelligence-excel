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

        private readonly Label lblStatistics;
        private readonly Label lblPercentiles;

        private readonly Panel histogramPanel;
        private readonly Panel tornadoPanel;

        private readonly TextBox txtTarget;
        private double? currentTarget;

        private readonly Label lblProbabilityBelow;
        private readonly Label lblProbabilityAbove;


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
            // SUMMARY HEADER
            // =====================================================

            Label lblSummaryHeader =
                new Label
                {
                    Text =
                        "Summary Statistics",

                    Left =
                        30,

                    Top =
                        125,

                    Width =
                        220,

                    Height =
                        25,

                    Font =
                        new Font(
                            "Segoe UI",
                            11,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblSummaryHeader);


            // =====================================================
            // SUMMARY VALUES
            // =====================================================

            lblStatistics =
                new Label
                {
                    Left =
                        30,

                    Top =
                        160,

                    Width =
                        220,

                    Height =
                        330,

                    Font =
                        new Font(
                            "Segoe UI",
                            10)
                };


            Controls.Add(
                lblStatistics);


            // =====================================================
            // PERCENTILE HEADER
            // =====================================================

            Label lblPercentileHeader =
                new Label
                {
                    Text =
                        "Percentiles",

                    Left =
                        270,

                    Top =
                        125,

                    Width =
                        180,

                    Height =
                        25,

                    Font =
                        new Font(
                            "Segoe UI",
                            11,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblPercentileHeader);


            lblPercentiles =
                new Label
                {
                    Left =
                        270,

                    Top =
                        160,

                    Width =
                        200,

                    Height =
                        260,

                    Font =
                        new Font(
                            "Segoe UI",
                            10)
                };


            Controls.Add(
                lblPercentiles);


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
                currentTarget = null;
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


            lblProbabilityBelow =
                new Label
                {
                    Left =
                        30,

                    Top =
                        575,

                    Width =
                        430,

                    Height =
                        30,

                    Font =
                        new Font(
                            "Segoe UI",
                            10,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblProbabilityBelow);


            lblProbabilityAbove =
                new Label
                {
                    Left =
                        30,

                    Top =
                        610,

                    Width =
                        430,

                    Height =
                        30,

                    Font =
                        new Font(
                            "Segoe UI",
                            10,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblProbabilityAbove);


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

            cmbForecast.SelectedIndex =
                0;


            RefreshForecastDisplay();
        }


        // =========================================================
        // FORECAST SELECTION CHANGED
        // =========================================================

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
            lblStatistics.Text =
                $"Trials\n" +
                $"{currentForecastResult.Trials:N0}\n\n" +

                $"Mean\n" +
                $"{FormatValue(currentForecastResult.Mean)}\n\n" +

                $"Std Deviation\n" +
                $"{FormatValue(currentForecastResult.StandardDeviation)}\n\n" +

                $"Minimum\n" +
                $"{FormatValue(currentForecastResult.Minimum)}\n\n" +

                $"Maximum\n" +
                $"{FormatValue(currentForecastResult.Maximum)}";


            lblPercentiles.Text =
                $"P10\n" +
                $"{FormatValue(currentForecastResult.P10)}\n\n" +

                $"P50\n" +
                $"{FormatValue(currentForecastResult.P50)}\n\n" +

                $"P80\n" +
                $"{FormatValue(currentForecastResult.P80)}\n\n" +

                $"P90\n" +
                $"{FormatValue(currentForecastResult.P90)}";


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
            currentTarget = double.IsFinite(target) ? target : null;

            double below =
                currentForecastResult
                    .ProbabilityLessThanOrEqual(
                        target);


            double above =
                currentForecastResult
                    .ProbabilityGreaterThan(
                        target);


            lblProbabilityBelow.Text =
                $"Probability ≤ {FormatValue(target)}: " +
                $"{below:P1}";


            lblProbabilityAbove.Text =
                $"Probability > {FormatValue(target)}: " +
                $"{above:P1}";
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

            int topMargin =
                20;

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


            DrawVerticalMarker(
                graphics,
                currentForecastResult.P50,
                min,
                max,
                leftMargin,
                topMargin,
                chartWidth,
                chartHeight,
                "P50");


            DrawVerticalMarker(
                graphics,
                currentForecastResult.P80,
                min,
                max,
                leftMargin,
                topMargin,
                chartWidth,
                chartHeight,
                "P80");


            DrawXAxisLabel(
                graphics,
                min,
                min,
                max,
                leftMargin,
                chartWidth,
                topMargin +
                    chartHeight +
                    8);


            DrawXAxisLabel(
                graphics,
                currentForecastResult.P50,
                min,
                max,
                leftMargin,
                chartWidth,
                topMargin +
                    chartHeight +
                    8);


            DrawXAxisLabel(
                graphics,
                currentForecastResult.P80,
                min,
                max,
                leftMargin,
                chartWidth,
                topMargin +
                    chartHeight +
                    8);


            DrawXAxisLabel(
                graphics,
                max,
                min,
                max,
                leftMargin,
                chartWidth,
                topMargin +
                    chartHeight +
                    8);


            DrawTargetMarker(graphics, area, min, max, leftMargin,
                topMargin, chartWidth, chartHeight);

            graphics.ResetClip();
        }

        private void DrawTargetMarker(Graphics graphics, Rectangle area,
            double min, double max, int leftMargin, int topMargin,
            int chartWidth, int chartHeight)
        {
            ChartValuePosition position = ChartValuePosition.Calculate(currentTarget, min, max);
            if (position.Range == ChartValueRange.Invalid || chartWidth <= 0 || chartHeight <= 0)
                return;

            string label = $"Target: {FormatValue(currentTarget!.Value)}";
            float labelX = leftMargin;
            using Pen pen = new Pen(Color.DarkOrange, 2.5f);
            if (position.Range == ChartValueRange.InRange)
            {
                float x = ValueToX(currentTarget.Value, min, max, leftMargin, chartWidth);
                graphics.DrawLine(pen, x, topMargin, x, topMargin + chartHeight);
                graphics.DrawLine(pen, x - 4, topMargin, x + 4, topMargin);
                labelX = x + 5;
            }
            else
            {
                label += position.Range == ChartValueRange.BelowRange
                    ? " — below simulated range" : " — above simulated range";
            }

            using Font font = new Font("Segoe UI", 8, FontStyle.Bold);
            using Brush textBrush = new SolidBrush(SystemColors.ControlText);
            using Brush background = new SolidBrush(SystemColors.Window);
            using StringFormat format = new StringFormat
            {
                Trimming = StringTrimming.EllipsisCharacter,
                FormatFlags = StringFormatFlags.NoWrap
            };
            SizeF size = graphics.MeasureString(label, font);
            float width = Math.Min(size.Width + 4, Math.Max(0, area.Width - 4));
            float height = Math.Min(size.Height + 2, Math.Max(0, area.Height - 4));
            labelX = Math.Clamp(labelX, area.Left + 2, area.Right - 2 - width);
            float labelY = Math.Clamp(topMargin + font.GetHeight(graphics) + 8,
                area.Top + 2, area.Bottom - 2 - height);
            RectangleF bounds = new RectangleF(labelX, labelY, width, height);
            graphics.FillRectangle(background, bounds);
            graphics.DrawString(label, font, textBrush, bounds, format);
        }


        // =========================================================
        // HISTOGRAM MARKER
        // =========================================================

        private void DrawVerticalMarker(
            Graphics graphics,
            double value,
            double min,
            double max,
            int leftMargin,
            int topMargin,
            int chartWidth,
            int chartHeight,
            string label)
        {
            float x =
                ValueToX(
                    value,
                    min,
                    max,
                    leftMargin,
                    chartWidth);


            using Pen markerPen =
                new Pen(
                    SystemColors.ControlDarkDark,
                    1.5f);


            markerPen.DashStyle =
                DashStyle.Dash;


            graphics.DrawLine(
                markerPen,
                x,
                topMargin,
                x,
                topMargin +
                    chartHeight);


            using Font markerFont =
                new Font(
                    "Segoe UI",
                    8,
                    FontStyle.Bold);


            using Brush markerTextBrush =
                new SolidBrush(
                    SystemColors.ControlText);


            graphics.DrawString(
                label,
                markerFont,
                markerTextBrush,
                x + 3,
                topMargin + 3);
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
