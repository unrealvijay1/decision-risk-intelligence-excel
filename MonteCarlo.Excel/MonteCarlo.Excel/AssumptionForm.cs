using System;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using MonteCarlo.Core;

namespace MonteCarlo.Excel
{
    public class AssumptionForm : Form
    {
        private readonly TextBox txtName;

        private readonly ComboBox cmbDistribution;

        private readonly Label lblParameter1;
        private readonly Label lblParameter2;
        private readonly Label lblParameter3;
        private readonly Label lblParameter4;

        private readonly TextBox txtParameter1;
        private readonly TextBox txtParameter2;
        private readonly TextBox txtParameter3;
        private readonly TextBox txtParameter4;
        private readonly Panel previewPanel;
        private MonteCarlo.Core.DistributionPreviewResult? preview;
        private readonly Panel tablePanel;
        private readonly DataGridView probabilityGrid;
        private readonly Label probabilityTotal;
        public DiscreteTable? ProbabilityTable { get; private set; }
        private DistributionKind SelectedKind => Enum.Parse<DistributionKind>((cmbDistribution.SelectedItem?.ToString() ?? "PERT").Replace(" ", ""), true);


        public string AssumptionName { get; private set; } = "";

        public DistributionType Distribution { get; private set; }

        public double Parameter1 { get; private set; }

        public double Parameter2 { get; private set; }

        public double Parameter3 { get; private set; }

        public double Parameter4 { get; private set; }


        // =========================================================
        // NEW ASSUMPTION
        // =========================================================

        public AssumptionForm(
            string cellAddress)
            : this(
                cellAddress,
                null)
        {
        }


        // =========================================================
        // NEW / EXISTING ASSUMPTION
        // =========================================================

        public AssumptionForm(
            string cellAddress,
            AssumptionDefinition? existingAssumption)
        {
            Text =
                existingAssumption == null
                    ? $"Define Assumption - {cellAddress}"
                    : $"Edit Assumption - {cellAddress}";


            Width =
                540;

            Height =
                700;

            StartPosition =
                FormStartPosition.CenterScreen;

            FormBorderStyle =
                FormBorderStyle.Sizable;

            MinimumSize = new Size(500, 610);
            AutoScaleMode = AutoScaleMode.Dpi;

            MaximizeBox =
                false;

            MinimizeBox =
                false;


            // =====================================================
            // ASSUMPTION NAME
            // =====================================================

            Label lblName =
                new Label
                {
                    Text =
                        "Assumption Name",

                    Left =
                        20,

                    Top =
                        30,

                    Width =
                        140
                };


            txtName =
                new TextBox
                {
                    Left =
                        170,

                    Top =
                        25,

                    Width =
                        250
                };


            Controls.Add(
                lblName);

            Controls.Add(
                txtName);


            // =====================================================
            // DISTRIBUTION
            // =====================================================

            Label lblDistribution =
                new Label
                {
                    Text =
                        "Distribution",

                    Left =
                        20,

                    Top =
                        80,

                    Width =
                        140
                };


            cmbDistribution =
                new ComboBox
                {
                    Left =
                        170,

                    Top =
                        75,

                    Width =
                        250,

                    DropDownStyle =
                        ComboBoxStyle.DropDownList
                };


            cmbDistribution.Items.Add(
                "PERT");

            cmbDistribution.Items.Add(
                "Triangular");

            cmbDistribution.Items.Add(
                "Normal");

            cmbDistribution.Items.Add(
                "Uniform");

            cmbDistribution.Items.Add(
                "Lognormal");

            cmbDistribution.Items.Add(
                "Beta");


            cmbDistribution.Items.AddRange(new object[] { "Exponential", "Poisson", "Binomial", "Discrete", "Weibull", "Gamma", "Bernoulli", "Truncated Normal" });
            Controls.Add(
                lblDistribution);

            Controls.Add(
                cmbDistribution);


            // =====================================================
            // PARAMETER 1
            // =====================================================

            lblParameter1 =
                new Label
                {
                    Left =
                        20,

                    Top =
                        135,

                    Width =
                        140
                };


            txtParameter1 =
                new TextBox
                {
                    Left =
                        170,

                    Top =
                        130,

                    Width =
                        250
                };


            Controls.Add(
                lblParameter1);

            Controls.Add(
                txtParameter1);


            // =====================================================
            // PARAMETER 2
            // =====================================================

            lblParameter2 =
                new Label
                {
                    Left =
                        20,

                    Top =
                        180,

                    Width =
                        140
                };


            txtParameter2 =
                new TextBox
                {
                    Left =
                        170,

                    Top =
                        175,

                    Width =
                        250
                };


            Controls.Add(
                lblParameter2);

            Controls.Add(
                txtParameter2);


            // =====================================================
            // PARAMETER 3
            // =====================================================

            lblParameter3 =
                new Label
                {
                    Left =
                        20,

                    Top =
                        225,

                    Width =
                        140
                };


            txtParameter3 =
                new TextBox
                {
                    Left =
                        170,

                    Top =
                        220,

                    Width =
                        250
                };


            Controls.Add(
                lblParameter3);

            Controls.Add(
                txtParameter3);


            // =====================================================
            // PARAMETER 4
            // =====================================================

            lblParameter4 =
                new Label
                {
                    Left =
                        20,

                    Top =
                        270,

                    Width =
                        140
                };


            txtParameter4 =
                new TextBox
                {
                    Left =
                        170,

                    Top =
                        265,

                    Width =
                        250
                };


            Controls.Add(
                lblParameter4);

            Controls.Add(
                txtParameter4);


            // =====================================================
            // FIT FROM DATA
            // =====================================================

            Button btnFitFromData =
                new Button
                {
                    Text =
                        "Fit from Selected Data",

                    Left =
                        170,

                    Top =
                        325,

                    Width =
                        250,

                    Height =
                        32
                };


            btnFitFromData.Click +=
                BtnFitFromData_Click;


            Controls.Add(
                btnFitFromData);

            var fittingHelp = new Label { Text = "Fitting: Normal, Lognormal, Uniform, Triangular only.", Left = 20, Top = 360, AutoSize = true };
            Controls.Add(fittingHelp);
            Label previewHeading = new Label
            {
                Text = "Distribution Preview",
                Left = 20,
                Top = 375,
                Width = 300,
                Font = new Font(Font, FontStyle.Bold)
            };
            Controls.Add(previewHeading);

            previewPanel = new PreviewPanel
            {
                Left = 20,
                Top = 400,
                Width = ClientSize.Width - 40,
                Height = ClientSize.Height - 485,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom |
                         AnchorStyles.Left | AnchorStyles.Right,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };
            previewPanel.Paint += PreviewPanel_Paint;
            Controls.Add(previewPanel);


            // =====================================================
            // BUTTONS
            // =====================================================

            Button btnOK =
                new Button
                {
                    Text =
                        "OK",

                    Left =
                        255,

                    Top =
                        ClientSize.Height - 60,

                    Width =
                        75,
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right
                };


            Button btnCancel =
                new Button
                {
                    Text =
                        "Cancel",

                    Left =
                        345,

                    Top =
                        ClientSize.Height - 60,

                    Width =
                        75,
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right
                };


            tablePanel = new Panel { Left = 20, Top = 120, Width = ClientSize.Width - 40, Height = 192,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, Visible = false };
            probabilityGrid = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Name = "ProbabilityTable" };
            probabilityGrid.Columns.Add("Outcome", "Outcome"); probabilityGrid.Columns.Add("Probability", "Probability");
            probabilityGrid.Rows.Add(10d, .2); probabilityGrid.Rows.Add(20d, .5); probabilityGrid.Rows.Add(30d, .3);
            var tableFooter = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 36, WrapContents = false };
            var addRow = new Button { Text = "Add Row", AutoSize = true };
            var deleteRow = new Button { Text = "Delete Row", AutoSize = true };
            probabilityTotal = new Label { Text = "Total: 1", AutoSize = true, Margin = new Padding(4, 8, 0, 0) };
            tableFooter.Controls.AddRange([addRow, deleteRow, probabilityTotal]);
            tablePanel.Controls.Add(probabilityGrid); tablePanel.Controls.Add(tableFooter); Controls.Add(tablePanel);
            addRow.Click += (_, _) => { if (probabilityGrid.Rows.Count < 100) probabilityGrid.Rows.Add(); UpdatePreview(); };
            deleteRow.Click += (_, _) => { if (probabilityGrid.CurrentRow is { } row) probabilityGrid.Rows.Remove(row); UpdatePreview(); };
            probabilityGrid.CellValueChanged += (_, _) => UpdatePreview();
            probabilityGrid.RowsRemoved += (_, _) => UpdatePreview();
            probabilityGrid.DataError += (_, e) => { e.ThrowException = false; };

            cmbDistribution.SelectedIndexChanged +=
                (_, _) =>
                {
                    UpdateParameterLabels();
                    if (SelectedKind >= DistributionKind.Exponential)
                    {
                        var defaults = DistributionCatalog.Defaults(SelectedKind);
                        TextBox[] fields = [txtParameter1, txtParameter2, txtParameter3, txtParameter4];
                        for (int i = 0; i < 4; i++) fields[i].Text = defaults[i].ToString(CultureInfo.CurrentCulture);
                    }
                    UpdatePreview();
                };

            txtParameter1.TextChanged += (_, _) => UpdatePreview();
            txtParameter2.TextChanged += (_, _) => UpdatePreview();
            txtParameter3.TextChanged += (_, _) => UpdatePreview();
            txtParameter4.TextChanged += (_, _) => UpdatePreview();


            btnOK.Click +=
                BtnOK_Click;


            btnCancel.Click +=
                (_, _) =>
                {
                    DialogResult =
                        DialogResult.Cancel;

                    Close();
                };


            Controls.Add(
                btnOK);

            Controls.Add(
                btnCancel);


            AcceptButton =
                btnOK;

            CancelButton =
                btnCancel;


            // =====================================================
            // INITIAL VALUES
            // =====================================================

            if (existingAssumption == null)
            {
                txtName.Text =
                    cellAddress;


                cmbDistribution.SelectedItem =
                    "PERT";


                UpdateParameterLabels();
            }
            else
            {
                LoadExistingAssumption(
                    existingAssumption);
            }

            UpdatePreview();
        }


        protected override void OnLayout(LayoutEventArgs e)
        {
            base.OnLayout(e);
            // Windows can constrain the scaled form to the working area. Keep the footer
            // and preview inside the actual client height after that constraint is applied.
            if (previewPanel == null || AcceptButton is not Button ok || CancelButton is not Button cancel) return;
            float scale = previewPanel.Top / 400f;
            ok.Top = cancel.Top = ClientSize.Height - (int)Math.Round(60 * scale);
            previewPanel.Height = Math.Max(40, ok.Top - (int)Math.Round(25 * scale) - previewPanel.Top);
        }

        private double[] ReadParameters()
        {
            var kind = SelectedKind;
            string[] names = DistributionCatalog.Parameters(kind);
            TextBox[] fields = [txtParameter1, txtParameter2, txtParameter3, txtParameter4];
            double[] values = new double[4];
            for (int i = 0; i < names.Length; i++)
                if (!double.TryParse(fields[i].Text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out values[i]))
                    throw new ArgumentException($"Enter a valid {names[i]}.");
            return values;
        }
        private DiscreteTable ReadTable()
        {
            var rows = new System.Collections.Generic.List<DiscreteOutcome>();
            foreach (DataGridViewRow row in probabilityGrid.Rows)
            {
                if (!double.TryParse(Convert.ToString(row.Cells[0].Value), out double outcome) ||
                    !double.TryParse(Convert.ToString(row.Cells[1].Value), out double probability))
                    throw new ArgumentException($"Enter numeric outcome and probability values in row {row.Index + 1}.");
                rows.Add(new(outcome, probability));
            }
            return new DiscreteTable(rows);
        }
        private void UpdatePreview()
        {
            if (previewPanel == null || probabilityGrid == null) return;
            try
            {
                var values = ReadParameters();
                if (SelectedKind == DistributionKind.Discrete)
                {
                    double total = probabilityGrid.Rows.Cast<DataGridViewRow>().Sum(r => double.TryParse(Convert.ToString(r.Cells[1].Value), out double p) ? p : 0);
                    probabilityTotal.Text = $"Total: {total:G8} (must be 1)";
                }
                var table = SelectedKind == DistributionKind.Discrete ? ReadTable() : null;
                preview = DistributionPreview.Generate(SelectedKind, values[0], values[1], values[2], values[3], table);
            }
            catch (ArgumentException ex) { preview = new() { ValidationMessage = ex.Message }; }
            previewPanel.Invalidate();
        }

        private void PreviewPanel_Paint(object? sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle bounds = previewPanel.ClientRectangle;
            if (preview == null || !preview.IsValid)
            {
                string message = preview?.ValidationMessage ?? "Enter distribution parameters.";
                TextRenderer.DrawText(g, message, Font,
                    new Rectangle(12, 12, Math.Max(1, bounds.Width - 24),
                        Math.Max(1, bounds.Height - 24)), Color.DarkRed,
                    TextFormatFlags.WordBreak);
                return;
            }

            float left = 54, right = bounds.Width - 18;
            float top = 14, bottom = bounds.Height - 49;
            if (right <= left || bottom <= top) return;

            using var axisPen = new Pen(Color.Gray, 1);
            g.DrawLine(axisPen, left, bottom, right, bottom);
            g.DrawLine(axisPen, left, top, left, bottom);

            double maxDensity = preview.Points.Max(p => p.Density);
            if (maxDensity <= 0) return;

            // Visual clipping affects only drawing, never Core density values.
            double[] sorted = preview.Points.Select(p => p.Density).OrderBy(v => v).ToArray();
            double visualMax = preview.IsDiscrete ? maxDensity : sorted[(int)(0.99 * (sorted.Length - 1))];
            if (visualMax <= 0) visualMax = maxDensity;
            var curve = new PointF[preview.Points.Count];
            for (int i = 0; i < curve.Length; i++)
            {
                var point = preview.Points[i];
                double fraction = (point.X - preview.MinimumX) /
                    (preview.MaximumX - preview.MinimumX);
                curve[i] = new PointF(
                    left + (float)(fraction * (right - left)),
                    bottom - (float)(Math.Min(point.Density, visualMax) /
                        visualMax * (bottom - top)));
            }
            using var curvePen = new Pen(Color.SteelBlue, 2.2f);
            if (preview.IsDiscrete)
            {
                float gap = curve.Length > 1 ? curve.Zip(curve.Skip(1), (a, b) => b.X - a.X).Min() : 12;
                float width = Math.Max(1, Math.Min(16, gap * .7f));
                using var fill = new SolidBrush(Color.SteelBlue);
                for (int i = 0; i < curve.Length; i++)
                    g.FillRectangle(fill, curve[i].X - width / 2, curve[i].Y, width, Math.Max(0, bottom - curve[i].Y));
                TextRenderer.DrawText(g, $"PMF max {maxDensity:G4}", Font, new Point((int)left, (int)top), Color.DimGray);
            }
            else g.DrawLines(curvePen, curve);

            string low = preview.MinimumX.ToString("G4", CultureInfo.CurrentCulture);
            string high = preview.MaximumX.ToString("G4", CultureInfo.CurrentCulture);
            TextRenderer.DrawText(g, low, Font,
                new Point((int)left, (int)bottom + 5), Color.DimGray);
            var highSize = TextRenderer.MeasureText(high, Font);
            TextRenderer.DrawText(g, high, Font,
                new Point((int)right - highSize.Width, (int)bottom + 5), Color.DimGray);
            TextRenderer.DrawText(g, preview.RangeNote ?? $"X range: {low} to {high}", Font,
                new Point((int)left, bounds.Height - 21), Color.DimGray);
        }

        private sealed class PreviewPanel : Panel
        {
            public PreviewPanel() => DoubleBuffered = true;
        }

        // =========================================================
        // FIT FROM SELECTED DATA
        // =========================================================

        private void BtnFitFromData_Click(
            object? sender,
            EventArgs e)
        {
            try
            {
                dynamic excelApp =
                    ExcelDna.Integration
                        .ExcelDnaUtil
                        .Application;


                Hide();


                dynamic selectedRange =
                    null;


                try
                {
                    selectedRange =
                        excelApp.InputBox(
                            Prompt:
                                "Select the historical data range.\n\n" +
                                "You can switch to another worksheet.",
                            Title:
                                "Monte Carlo - Select Historical Data",
                            Type:
                                8);
                }
                catch
                {
                    Show();

                    return;
                }


                if (selectedRange == null)
                {
                    Show();

                    return;
                }


                string rangeAddress;


                double[] data =
                    ExcelDataReader
                        .ReadNumericDataFromRange(
                            selectedRange,
                            out rangeAddress);


                Show();


                using DistributionFitForm fitForm =
                    new DistributionFitForm(
                        rangeAddress,
                        data);


                if (fitForm.ShowDialog(this)
                    != DialogResult.OK)
                {
                    return;
                }


                if (fitForm.SelectedFit == null)
                {
                    return;
                }


                ApplyFittedDistribution(
                    fitForm.SelectedFit);
            }
            catch (Exception ex)
            {
                Show();


                MessageBox.Show(
                    ex.Message,
                    "Distribution Fitting",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }


        // =========================================================
        // APPLY FIT RESULT
        //
        // Beta fitting is not wired yet. This continues to support
        // the currently fitted distributions.
        // =========================================================

        private void ApplyFittedDistribution(
            MonteCarlo.Core.DistributionFitResult fit)
        {
            switch (fit.Distribution)
            {
                case MonteCarlo.Core
                    .FittedDistributionType.Normal:

                    cmbDistribution.SelectedItem =
                        "Normal";

                    txtParameter1.Text =
                        fit.Parameter1.ToString(
                            "0.########");

                    txtParameter2.Text =
                        fit.Parameter2.ToString(
                            "0.########");

                    txtParameter3.Text =
                        "";

                    txtParameter4.Text =
                        "";

                    break;


                case MonteCarlo.Core
                    .FittedDistributionType.Triangular:

                    cmbDistribution.SelectedItem =
                        "Triangular";

                    txtParameter1.Text =
                        fit.Parameter1.ToString(
                            "0.########");

                    txtParameter2.Text =
                        fit.Parameter2.ToString(
                            "0.########");

                    txtParameter3.Text =
                        fit.Parameter3.ToString(
                            "0.########");

                    txtParameter4.Text =
                        "";

                    break;


                case MonteCarlo.Core
                    .FittedDistributionType.Uniform:

                    cmbDistribution.SelectedItem =
                        "Uniform";

                    txtParameter1.Text =
                        fit.Parameter1.ToString(
                            "0.########");

                    txtParameter2.Text =
                        fit.Parameter2.ToString(
                            "0.########");

                    txtParameter3.Text =
                        "";

                    txtParameter4.Text =
                        "";

                    break;


                case MonteCarlo.Core
                    .FittedDistributionType.Lognormal:

                    cmbDistribution.SelectedItem =
                        "Lognormal";

                    txtParameter1.Text =
                        fit.Parameter1.ToString(
                            "0.########");

                    txtParameter2.Text =
                        fit.Parameter2.ToString(
                            "0.########");

                    txtParameter3.Text =
                        "";

                    txtParameter4.Text =
                        "";

                    break;
            }


            UpdateParameterLabels();


            MessageBox.Show(
                $"Selected distribution: {fit.Distribution}\n\n" +
                $"{fit.ParameterDescription}\n\n" +
                $"AIC: {fit.AIC:N3}\n" +
                $"KS: {fit.KSStatistic:0.0000}",
                "Distribution Fitting",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }


        // =========================================================
        // LOAD EXISTING ASSUMPTION
        // =========================================================

        private void LoadExistingAssumption(
            AssumptionDefinition assumption)
        {
            txtName.Text =
                string.IsNullOrWhiteSpace(
                    assumption.Name)
                    ? $"{assumption.SheetName}!" +
                      $"{assumption.CellAddress}"
                    : assumption.Name;


            cmbDistribution.SelectedItem = DistributionCatalog.Name((DistributionKind)assumption.Distribution);
            if (assumption.ProbabilityTable != null)
            {
                probabilityGrid.Rows.Clear();
                foreach (var row in assumption.ProbabilityTable.Outcomes) probabilityGrid.Rows.Add(row.Outcome, row.Probability);
            }

            UpdateParameterLabels();


            txtParameter1.Text =
                assumption.Parameter1
                    .ToString();


            txtParameter2.Text =
                assumption.Parameter2
                    .ToString();


            txtParameter3.Text =
                assumption.Parameter3
                    .ToString();


            txtParameter4.Text =
                assumption.Parameter4
                    .ToString();
        }


        // =========================================================
        // UPDATE PARAMETER LABELS
        // =========================================================

        private void UpdateParameterLabels()
        {
            string[] names = DistributionCatalog.Parameters(SelectedKind);
            Label[] labels = [lblParameter1, lblParameter2, lblParameter3, lblParameter4];
            TextBox[] fields = [txtParameter1, txtParameter2, txtParameter3, txtParameter4];
            for (int i = 0; i < 4; i++)
            {
                labels[i].Visible = fields[i].Visible = i < names.Length;
                if (i < names.Length) labels[i].Text = names[i];
            }
            tablePanel.Visible = SelectedKind == DistributionKind.Discrete;
        }

        private void BtnOK_Click(object? sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtName.Text)) throw new ArgumentException("Enter an assumption name.");
                probabilityGrid.EndEdit();
                var values = ReadParameters();
                var table = SelectedKind == DistributionKind.Discrete ? ReadTable() : null;
                if (DistributionPreview.Validate(SelectedKind, values[0], values[1], values[2], values[3], table) is { } error)
                    throw new ArgumentException(error);
                AssumptionName = txtName.Text.Trim(); Distribution = (DistributionType)SelectedKind;
                Parameter1 = values[0]; Parameter2 = values[1]; Parameter3 = values[2]; Parameter4 = values[3]; ProbabilityTable = table;
                DialogResult = DialogResult.OK; Close();
            }
            catch (ArgumentException ex) { MessageBox.Show(ex.Message, "Monte Carlo"); }
        }
    }
}
