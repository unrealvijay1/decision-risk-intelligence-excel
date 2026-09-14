using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace MonteCarlo.Excel
{
    public class DistributionFitForm : Form
    {
        private readonly ListView listView;
        private readonly Panel chartPanel;
        private readonly Label lblSelectedFit;

        private readonly double[] historicalData;

        private readonly List<
            MonteCarlo.Core.DistributionFitResult> fitResults;


        public MonteCarlo.Core.DistributionFitResult?
            SelectedFit
        { get; private set; }


        // =========================================================
        // CONSTRUCTOR
        // =========================================================

        public DistributionFitForm(
            string rangeAddress,
            double[] data)
        {
            if (data == null)
            {
                throw new ArgumentNullException(
                    nameof(data));
            }

            if (data.Length < 5)
            {
                throw new ArgumentException(
                    "At least 5 numeric observations are required.");
            }


            historicalData =
                data.ToArray();


            Text =
                "Fit Distribution";

            Width =
                1250;

            Height =
                760;

            StartPosition =
                FormStartPosition.CenterParent;

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
                        "Distribution Fitting",

                    Left =
                        20,

                    Top =
                        15,

                    Width =
                        400,

                    Height =
                        35,

                    Font =
                        new Font(
                            "Segoe UI",
                            16,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblTitle);


            // =====================================================
            // DATA INFORMATION
            // =====================================================

            Label lblRange =
                new Label
                {
                    Text =
                        $"Historical Data: {rangeAddress}",

                    Left =
                        20,

                    Top =
                        58,

                    Width =
                        750,

                    Height =
                        25
                };


            Controls.Add(
                lblRange);


            Label lblCount =
                new Label
                {
                    Text =
                        $"Observations: {historicalData.Length:N0}",

                    Left =
                        20,

                    Top =
                        85,

                    Width =
                        300,

                    Height =
                        25
                };


            Controls.Add(
                lblCount);


            Label lblInfo =
                new Label
                {
                    Text =
                        "Select a distribution below to compare its fitted curve " +
                        "with the historical data. Lower AIC and KS generally indicate better fit.",

                    Left =
                        20,

                    Top =
                        115,

                    Width =
                        1150,

                    Height =
                        35
                };


            Controls.Add(
                lblInfo);


            // =====================================================
            // TABLE HEADER
            // =====================================================

            Label lblRanking =
                new Label
                {
                    Text =
                        "Candidate Distributions",

                    Left =
                        20,

                    Top =
                        155,

                    Width =
                        300,

                    Height =
                        25,

                    Font =
                        new Font(
                            "Segoe UI",
                            11,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblRanking);


            // =====================================================
            // FIT RESULTS TABLE
            // =====================================================

            listView =
                new ListView
                {
                    Left =
                        20,

                    Top =
                        185,

                    Width =
                        575,

                    Height =
                        430,

                    View =
                        View.Details,

                    FullRowSelect =
                        true,

                    GridLines =
                        true,

                    MultiSelect =
                        false,

                    HideSelection =
                        false
                };


            listView.Columns.Add(
                "Rank",
                50);

            listView.Columns.Add(
                "Distribution",
                100);

            listView.Columns.Add(
                "AIC",
                105);

            listView.Columns.Add(
                "ΔAIC",
                90);

            listView.Columns.Add(
                "KS",
                80);

            listView.Columns.Add(
                "Parameters",
                220);


            listView.SelectedIndexChanged +=
                ListView_SelectedIndexChanged;


            Controls.Add(
                listView);


            // =====================================================
            // CHART HEADER
            // =====================================================

            Label lblChartHeader =
                new Label
                {
                    Text =
                        "Historical Data vs Fitted Distribution",

                    Left =
                        625,

                    Top =
                        155,

                    Width =
                        450,

                    Height =
                        25,

                    Font =
                        new Font(
                            "Segoe UI",
                            11,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblChartHeader);


            // =====================================================
            // CHART
            // =====================================================

            chartPanel =
                new Panel
                {
                    Left =
                        625,

                    Top =
                        185,

                    Width =
                        585,

                    Height =
                        370,

                    BorderStyle =
                        BorderStyle.FixedSingle,

                    BackColor =
                        SystemColors.Window
                };


            chartPanel.Paint +=
                ChartPanel_Paint;


            Controls.Add(
                chartPanel);


            // =====================================================
            // SELECTED FIT INFORMATION
            // =====================================================

            lblSelectedFit =
                new Label
                {
                    Left =
                        625,

                    Top =
                        570,

                    Width =
                        585,

                    Height =
                        80,

                    Font =
                        new Font(
                            "Segoe UI",
                            9)
                };


            Controls.Add(
                lblSelectedFit);


            // =====================================================
            // BUTTONS
            // =====================================================

            Button btnUseSelected =
                new Button
                {
                    Text =
                        "Use Selected Distribution",

                    Left =
                        900,

                    Top =
                        665,

                    Width =
                        205,

                    Height =
                        32
                };


            Button btnCancel =
                new Button
                {
                    Text =
                        "Cancel",

                    Left =
                        1120,

                    Top =
                        665,

                    Width =
                        90,

                    Height =
                        32
                };


            btnUseSelected.Click +=
                BtnUseSelected_Click;


            btnCancel.Click +=
                (_, _) =>
                {
                    DialogResult =
                        DialogResult.Cancel;

                    Close();
                };


            Controls.Add(
                btnUseSelected);

            Controls.Add(
                btnCancel);


            AcceptButton =
                btnUseSelected;

            CancelButton =
                btnCancel;


            // =====================================================
            // FIT DISTRIBUTIONS
            // =====================================================

            try
            {
                fitResults =
                    MonteCarlo.Core
                        .DistributionFitter
                        .FitAll(
                            historicalData);


                LoadFitResults();
            }
            catch (Exception ex)
            {
                fitResults =
                    new List<
                        MonteCarlo.Core.DistributionFitResult>();


                MessageBox.Show(
                    ex.Message,
                    "Distribution Fitting",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }


        // =========================================================
        // LOAD FIT RESULTS
        // =========================================================

        private void LoadFitResults()
        {
            listView.Items.Clear();


            if (fitResults.Count == 0)
            {
                return;
            }


            double bestAic =
                fitResults.Min(
                    x => x.AIC);


            foreach (
                MonteCarlo.Core.DistributionFitResult fit
                in fitResults)
            {
                double deltaAic =
                    fit.AIC -
                    bestAic;


                ListViewItem item =
                    new ListViewItem(
                        fit.Rank.ToString());


                item.SubItems.Add(
                    fit.Distribution.ToString());


                item.SubItems.Add(
                    fit.AIC.ToString(
                        "N2"));


                item.SubItems.Add(
                    deltaAic.ToString(
                        "N2"));


                item.SubItems.Add(
                    fit.KSStatistic.ToString(
                        "0.0000"));


                item.SubItems.Add(
                    GetShortParameterDescription(
                        fit));


                item.Tag =
                    fit;


                listView.Items.Add(
                    item);
            }


            if (listView.Items.Count > 0)
            {
                listView.Items[0].Selected =
                    true;

                listView.Items[0].Focused =
                    true;

                listView.EnsureVisible(
                    0);
            }
        }


        // =========================================================
        // TABLE SELECTION CHANGED
        // =========================================================

        private void ListView_SelectedIndexChanged(
            object? sender,
            EventArgs e)
        {
            if (listView.SelectedItems.Count == 0)
            {
                return;
            }


            if (
                listView.SelectedItems[0].Tag
                is not MonteCarlo.Core.DistributionFitResult fit)
            {
                return;
            }


            SelectedFit =
                fit;


            UpdateSelectedFitDescription(
                fit);


            chartPanel.Invalidate();
        }


        // =========================================================
        // SELECTED FIT DESCRIPTION
        // =========================================================

        private void UpdateSelectedFitDescription(
            MonteCarlo.Core.DistributionFitResult fit)
        {
            string interpretation =
                GetInterpretation(
                    fit);


            lblSelectedFit.Text =
                $"Selected: {fit.Distribution}\n" +
                $"{interpretation}\n" +
                $"AIC = {fit.AIC:N2}    " +
                $"KS = {fit.KSStatistic:0.0000}";
        }


        // =========================================================
        // PARAMETER DESCRIPTION
        // =========================================================

        private static string GetShortParameterDescription(
            MonteCarlo.Core.DistributionFitResult fit)
        {
            switch (fit.Distribution)
            {
                case MonteCarlo.Core
                    .FittedDistributionType.Normal:

                    return
                        $"μ={fit.Parameter1:N1}, " +
                        $"σ={fit.Parameter2:N1}";


                case MonteCarlo.Core
                    .FittedDistributionType.Lognormal:

                    return
                        $"log μ={fit.Parameter1:N3}, " +
                        $"σ={fit.Parameter2:N3}";


                case MonteCarlo.Core
                    .FittedDistributionType.Uniform:

                    return
                        $"{fit.Parameter1:N1} – " +
                        $"{fit.Parameter2:N1}";


                case MonteCarlo.Core
                    .FittedDistributionType.Triangular:

                    return
                        $"{fit.Parameter1:N1}, " +
                        $"{fit.Parameter2:N1}, " +
                        $"{fit.Parameter3:N1}";


                default:

                    return "";
            }
        }


        // =========================================================
        // ORIGINAL SCALE INTERPRETATION
        // =========================================================

        private static string GetInterpretation(
            MonteCarlo.Core.DistributionFitResult fit)
        {
            switch (fit.Distribution)
            {
                case MonteCarlo.Core
                    .FittedDistributionType.Normal:

                    return
                        $"Mean ≈ {fit.Parameter1:N2}, " +
                        $"SD ≈ {fit.Parameter2:N2}";


                case MonteCarlo.Core
                    .FittedDistributionType.Lognormal:
                    {
                        double mu =
                            fit.Parameter1;


                        double sigma =
                            fit.Parameter2;


                        double median =
                            Math.Exp(
                                mu);


                        double mean =
                            Math.Exp(
                                mu +
                                sigma *
                                sigma /
                                2.0);


                        double variance =
                            (
                                Math.Exp(
                                    sigma *
                                    sigma)
                                -
                                1.0
                            )
                            *
                            Math.Exp(
                                2.0 *
                                mu
                                +
                                sigma *
                                sigma);


                        double sd =
                            Math.Sqrt(
                                variance);


                        return
                            $"Original scale: Mean ≈ {mean:N2}, " +
                            $"Median ≈ {median:N2}, " +
                            $"SD ≈ {sd:N2}";
                    }


                case MonteCarlo.Core
                    .FittedDistributionType.Uniform:

                    return
                        $"Range ≈ {fit.Parameter1:N2} to " +
                        $"{fit.Parameter2:N2}";


                case MonteCarlo.Core
                    .FittedDistributionType.Triangular:

                    return
                        $"Min ≈ {fit.Parameter1:N2}, " +
                        $"Mode ≈ {fit.Parameter2:N2}, " +
                        $"Max ≈ {fit.Parameter3:N2}";


                default:

                    return "";
            }
        }


        // =========================================================
        // CHART PAINT
        // =========================================================

        private void ChartPanel_Paint(
            object? sender,
            PaintEventArgs e)
        {
            DrawFitChart(
                e.Graphics,
                chartPanel.ClientRectangle);
        }


        // =========================================================
        // DRAW FIT CHART
        // =========================================================

        private void DrawFitChart(
            Graphics graphics,
            Rectangle area)
        {
            graphics.SmoothingMode =
                SmoothingMode.AntiAlias;


            graphics.Clear(
                SystemColors.Window);


            if (historicalData.Length == 0)
            {
                return;
            }


            double dataMin =
                historicalData.Min();


            double dataMax =
                historicalData.Max();


            if (dataMax <= dataMin)
            {
                return;
            }


            // Add a little space around the data.
            double range =
                dataMax -
                dataMin;


            double xMin =
                dataMin -
                range *
                0.05;


            double xMax =
                dataMax +
                range *
                0.05;


            // Lognormal must not use negative x.
            if (
                SelectedFit != null
                &&
                SelectedFit.Distribution ==
                    MonteCarlo.Core
                        .FittedDistributionType
                        .Lognormal)
            {
                xMin =
                    Math.Max(
                        0,
                        xMin);
            }


            int leftMargin =
                60;

            int rightMargin =
                25;

            int topMargin =
                25;

            int bottomMargin =
                50;


            int chartWidth =
                area.Width -
                leftMargin -
                rightMargin;


            int chartHeight =
                area.Height -
                topMargin -
                bottomMargin;


            if (
                chartWidth <= 0
                ||
                chartHeight <= 0)
            {
                return;
            }


            // =====================================================
            // HISTOGRAM
            // =====================================================

            int binCount =
                Math.Max(
                    8,
                    Math.Min(
                        25,
                        (int)Math.Sqrt(
                            historicalData.Length)));


            int[] bins =
                new int[binCount];


            double binWidth =
                (dataMax - dataMin)
                /
                binCount;


            if (binWidth <= 0)
            {
                return;
            }


            foreach (
                double value
                in historicalData)
            {
                int index =
                    (int)(
                        (value - dataMin)
                        /
                        binWidth);


                if (index >= binCount)
                {
                    index =
                        binCount - 1;
                }


                if (index < 0)
                {
                    index =
                        0;
                }


                bins[index]++;
            }


            int maxFrequency =
                bins.Max();


            // =====================================================
            // PDF MAXIMUM
            // =====================================================

            double maxPdf =
                0;


            if (SelectedFit != null)
            {
                const int pdfSteps =
                    300;


                for (
                    int i = 0;
                    i <= pdfSteps;
                    i++)
                {
                    double x =
                        xMin
                        +
                        (
                            xMax -
                            xMin
                        )
                        *
                        i
                        /
                        pdfSteps;


                    double pdf =
                        GetPdf(
                            SelectedFit,
                            x);


                    if (
                        !double.IsNaN(pdf)
                        &&
                        !double.IsInfinity(pdf))
                    {
                        maxPdf =
                            Math.Max(
                                maxPdf,
                                pdf);
                    }
                }
            }


            // =====================================================
            // AXES
            // =====================================================

            using Pen axisPen =
                new Pen(
                    SystemColors.ControlText);


            graphics.DrawLine(
                axisPen,
                leftMargin,
                topMargin,
                leftMargin,
                topMargin +
                    chartHeight);


            graphics.DrawLine(
                axisPen,
                leftMargin,
                topMargin +
                    chartHeight,
                leftMargin +
                    chartWidth,
                topMargin +
                    chartHeight);


            // =====================================================
            // HISTOGRAM BARS
            // =====================================================

            using Brush barBrush =
                new SolidBrush(
                    Color.FromArgb(
                        190,
                        SystemColors.Highlight));


            for (
                int i = 0;
                i < binCount;
                i++)
            {
                double binStart =
                    dataMin +
                    i *
                    binWidth;


                double binEnd =
                    binStart +
                    binWidth;


                float x1 =
                    MapX(
                        binStart,
                        xMin,
                        xMax,
                        leftMargin,
                        chartWidth);


                float x2 =
                    MapX(
                        binEnd,
                        xMin,
                        xMax,
                        leftMargin,
                        chartWidth);


                float barHeight =
                    maxFrequency == 0
                        ? 0
                        : (float)bins[i]
                          /
                          maxFrequency
                          *
                          chartHeight
                          *
                          0.90f;


                float y =
                    topMargin +
                    chartHeight -
                    barHeight;


                graphics.FillRectangle(
                    barBrush,
                    x1 + 1,
                    y,
                    Math.Max(
                        1,
                        x2 -
                        x1 -
                        2),
                    barHeight);
            }


            // =====================================================
            // FITTED PDF CURVE
            //
            // Scale PDF to histogram height.
            // =====================================================

            if (
                SelectedFit != null
                &&
                maxPdf > 0)
            {
                using Pen curvePen =
                    new Pen(
                        Color.DarkRed,
                        2.5f);


                PointF? previous =
                    null;


                const int curveSteps =
                    400;


                for (
                    int i = 0;
                    i <= curveSteps;
                    i++)
                {
                    double x =
                        xMin
                        +
                        (
                            xMax -
                            xMin
                        )
                        *
                        i
                        /
                        curveSteps;


                    double pdf =
                        GetPdf(
                            SelectedFit,
                            x);


                    if (
                        double.IsNaN(pdf)
                        ||
                        double.IsInfinity(pdf)
                        ||
                        pdf < 0)
                    {
                        previous =
                            null;

                        continue;
                    }


                    float px =
                        MapX(
                            x,
                            xMin,
                            xMax,
                            leftMargin,
                            chartWidth);


                    float normalized =
                        (float)(
                            pdf /
                            maxPdf);


                    float py =
                        topMargin +
                        chartHeight -
                        normalized
                        *
                        chartHeight
                        *
                        0.90f;


                    PointF current =
                        new PointF(
                            px,
                            py);


                    if (previous.HasValue)
                    {
                        graphics.DrawLine(
                            curvePen,
                            previous.Value,
                            current);
                    }


                    previous =
                        current;
                }
            }


            // =====================================================
            // AXIS LABELS
            // =====================================================

            using Font axisFont =
                new Font(
                    "Segoe UI",
                    8);


            using Brush textBrush =
                new SolidBrush(
                    SystemColors.ControlText);


            DrawAxisLabel(
                graphics,
                FormatCompactNumber(
                    dataMin),
                MapX(
                    dataMin,
                    xMin,
                    xMax,
                    leftMargin,
                    chartWidth),
                topMargin +
                    chartHeight +
                    8,
                axisFont,
                textBrush);


            double midpoint =
                (
                    dataMin +
                    dataMax
                )
                /
                2.0;


            DrawAxisLabel(
                graphics,
                FormatCompactNumber(
                    midpoint),
                MapX(
                    midpoint,
                    xMin,
                    xMax,
                    leftMargin,
                    chartWidth),
                topMargin +
                    chartHeight +
                    8,
                axisFont,
                textBrush);


            DrawAxisLabel(
                graphics,
                FormatCompactNumber(
                    dataMax),
                MapX(
                    dataMax,
                    xMin,
                    xMax,
                    leftMargin,
                    chartWidth),
                topMargin +
                    chartHeight +
                    8,
                axisFont,
                textBrush);


            // =====================================================
            // LEGEND
            // =====================================================

            if (SelectedFit != null)
            {
                using Brush legendBrush =
                    new SolidBrush(
                        Color.DarkRed);


                graphics.FillRectangle(
                    legendBrush,
                    leftMargin + 10,
                    topMargin + 8,
                    22,
                    3);


                graphics.DrawString(
                    $"{SelectedFit.Distribution} fitted curve",
                    axisFont,
                    textBrush,
                    leftMargin + 40,
                    topMargin + 1);
            }
        }


        // =========================================================
        // PDF
        // =========================================================

        private static double GetPdf(
            MonteCarlo.Core.DistributionFitResult fit,
            double x)
        {
            switch (fit.Distribution)
            {
                // -------------------------------------------------
                // NORMAL
                // -------------------------------------------------

                case MonteCarlo.Core
                    .FittedDistributionType.Normal:
                    {
                        double mean =
                            fit.Parameter1;


                        double sd =
                            fit.Parameter2;


                        if (sd <= 0)
                        {
                            return 0;
                        }


                        double z =
                            (x - mean)
                            /
                            sd;


                        return
                            Math.Exp(
                                -0.5 *
                                z *
                                z)
                            /
                            (
                                sd *
                                Math.Sqrt(
                                    2.0 *
                                    Math.PI)
                            );
                    }


                // -------------------------------------------------
                // LOGNORMAL
                // -------------------------------------------------

                case MonteCarlo.Core
                    .FittedDistributionType.Lognormal:
                    {
                        if (x <= 0)
                        {
                            return 0;
                        }


                        double mu =
                            fit.Parameter1;


                        double sigma =
                            fit.Parameter2;


                        if (sigma <= 0)
                        {
                            return 0;
                        }


                        double logX =
                            Math.Log(
                                x);


                        double z =
                            (logX - mu)
                            /
                            sigma;


                        return
                            Math.Exp(
                                -0.5 *
                                z *
                                z)
                            /
                            (
                                x *
                                sigma *
                                Math.Sqrt(
                                    2.0 *
                                    Math.PI)
                            );
                    }


                // -------------------------------------------------
                // UNIFORM
                // -------------------------------------------------

                case MonteCarlo.Core
                    .FittedDistributionType.Uniform:
                    {
                        double min =
                            fit.Parameter1;


                        double max =
                            fit.Parameter2;


                        if (
                            max <= min
                            ||
                            x < min
                            ||
                            x > max)
                        {
                            return 0;
                        }


                        return
                            1.0 /
                            (max - min);
                    }


                // -------------------------------------------------
                // TRIANGULAR
                // -------------------------------------------------

                case MonteCarlo.Core
                    .FittedDistributionType.Triangular:
                    {
                        double min =
                            fit.Parameter1;


                        double mode =
                            fit.Parameter2;


                        double max =
                            fit.Parameter3;


                        if (
                            max <= min
                            ||
                            mode < min
                            ||
                            mode > max
                            ||
                            x < min
                            ||
                            x > max)
                        {
                            return 0;
                        }


                        if (x == mode)
                        {
                            return
                                2.0 /
                                (max - min);
                        }


                        if (x < mode)
                        {
                            if (mode <= min)
                            {
                                return 0;
                            }


                            return
                                2.0 *
                                (x - min)
                                /
                                (
                                    (max - min)
                                    *
                                    (mode - min)
                                );
                        }


                        if (mode >= max)
                        {
                            return 0;
                        }


                        return
                            2.0 *
                            (max - x)
                            /
                            (
                                (max - min)
                                *
                                (max - mode)
                            );
                    }


                default:

                    return 0;
            }
        }


        // =========================================================
        // MAP X
        // =========================================================

        private static float MapX(
            double value,
            double min,
            double max,
            int left,
            int width)
        {
            if (max <= min)
            {
                return left;
            }


            double ratio =
                (value - min)
                /
                (max - min);


            ratio =
                Math.Max(
                    0,
                    Math.Min(
                        1,
                        ratio));


            return
                left +
                (float)(
                    ratio *
                    width);
        }


        // =========================================================
        // AXIS LABEL
        // =========================================================

        private static void DrawAxisLabel(
            Graphics graphics,
            string text,
            float x,
            float y,
            Font font,
            Brush brush)
        {
            SizeF size =
                graphics.MeasureString(
                    text,
                    font);


            graphics.DrawString(
                text,
                font,
                brush,
                x -
                    size.Width /
                    2,
                y);
        }


        // =========================================================
        // NUMBER FORMAT
        // =========================================================

        private static string FormatCompactNumber(
            double value)
        {
            double absolute =
                Math.Abs(
                    value);


            if (absolute >= 10000000)
            {
                return
                    (
                        value /
                        10000000.0
                    )
                    .ToString(
                        "0.0")
                    +
                    "Cr";
            }


            if (absolute >= 100000)
            {
                return
                    (
                        value /
                        100000.0
                    )
                    .ToString(
                        "0.0")
                    +
                    "L";
            }


            if (absolute >= 1000)
            {
                return
                    (
                        value /
                        1000.0
                    )
                    .ToString(
                        "0.0")
                    +
                    "K";
            }


            return
                value.ToString(
                    "0.##");
        }


        // =========================================================
        // USE SELECTED DISTRIBUTION
        // =========================================================

        private void BtnUseSelected_Click(
            object? sender,
            EventArgs e)
        {
            if (listView.SelectedItems.Count == 0)
            {
                MessageBox.Show(
                    "Select a fitted distribution first.",
                    "Distribution Fitting",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return;
            }


            if (
                listView.SelectedItems[0].Tag
                is not MonteCarlo.Core.DistributionFitResult fit)
            {
                MessageBox.Show(
                    "Unable to read the selected fit.",
                    "Distribution Fitting");

                return;
            }


            SelectedFit =
                fit;


            DialogResult =
                DialogResult.OK;


            Close();
        }
    }
}