using System;
using System.Windows.Forms;

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
                480;

            Height =
                540;

            StartPosition =
                FormStartPosition.CenterScreen;

            FormBorderStyle =
                FormBorderStyle.FixedDialog;

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
                        405,

                    Width =
                        75
                };


            Button btnCancel =
                new Button
                {
                    Text =
                        "Cancel",

                    Left =
                        345,

                    Top =
                        405,

                    Width =
                        75
                };


            cmbDistribution.SelectedIndexChanged +=
                (_, _) =>
                {
                    UpdateParameterLabels();
                };


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


            switch (assumption.Distribution)
            {
                case DistributionType.Normal:

                    cmbDistribution.SelectedItem =
                        "Normal";

                    break;


                case DistributionType.Triangular:

                    cmbDistribution.SelectedItem =
                        "Triangular";

                    break;


                case DistributionType.Pert:

                    cmbDistribution.SelectedItem =
                        "PERT";

                    break;


                case DistributionType.Uniform:

                    cmbDistribution.SelectedItem =
                        "Uniform";

                    break;


                case DistributionType.Lognormal:

                    cmbDistribution.SelectedItem =
                        "Lognormal";

                    break;


                case DistributionType.Beta:

                    cmbDistribution.SelectedItem =
                        "Beta";

                    break;


                default:

                    cmbDistribution.SelectedItem =
                        "PERT";

                    break;
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
            string distribution =
                cmbDistribution
                    .SelectedItem?
                    .ToString()
                ?? "PERT";


            // Hide optional fields by default.
            lblParameter3.Visible =
                false;

            txtParameter3.Visible =
                false;

            lblParameter4.Visible =
                false;

            txtParameter4.Visible =
                false;


            // -----------------------------------------------------
            // NORMAL
            // -----------------------------------------------------

            if (distribution == "Normal")
            {
                lblParameter1.Text =
                    "Mean";

                lblParameter2.Text =
                    "Std Deviation";

                return;
            }


            // -----------------------------------------------------
            // LOGNORMAL
            // -----------------------------------------------------

            if (distribution == "Lognormal")
            {
                lblParameter1.Text =
                    "Log Mean";

                lblParameter2.Text =
                    "Log Std Dev";

                return;
            }


            // -----------------------------------------------------
            // UNIFORM
            // -----------------------------------------------------

            if (distribution == "Uniform")
            {
                lblParameter1.Text =
                    "Minimum";

                lblParameter2.Text =
                    "Maximum";

                return;
            }


            // -----------------------------------------------------
            // BETA
            // -----------------------------------------------------

            if (distribution == "Beta")
            {
                lblParameter1.Text =
                    "Minimum";

                lblParameter2.Text =
                    "Maximum";

                lblParameter3.Text =
                    "Alpha";

                lblParameter4.Text =
                    "Beta";


                lblParameter3.Visible =
                    true;

                txtParameter3.Visible =
                    true;

                lblParameter4.Visible =
                    true;

                txtParameter4.Visible =
                    true;

                return;
            }


            // -----------------------------------------------------
            // PERT / TRIANGULAR
            // -----------------------------------------------------

            lblParameter1.Text =
                "Minimum";

            lblParameter2.Text =
                "Most Likely";

            lblParameter3.Text =
                "Maximum";


            lblParameter3.Visible =
                true;

            txtParameter3.Visible =
                true;
        }


        // =========================================================
        // OK
        // =========================================================

        private void BtnOK_Click(
            object? sender,
            EventArgs e)
        {
            string assumptionName =
                txtName.Text.Trim();


            if (string.IsNullOrWhiteSpace(
                    assumptionName))
            {
                MessageBox.Show(
                    "Enter an assumption name.",
                    "Monte Carlo");

                return;
            }


            string distribution =
                cmbDistribution
                    .SelectedItem?
                    .ToString()
                ?? "PERT";


            // -----------------------------------------------------
            // PARAMETER 1
            // -----------------------------------------------------

            if (!double.TryParse(
                    txtParameter1.Text,
                    out double parameter1))
            {
                MessageBox.Show(
                    "Enter a valid first parameter.");

                return;
            }


            // -----------------------------------------------------
            // PARAMETER 2
            // -----------------------------------------------------

            if (!double.TryParse(
                    txtParameter2.Text,
                    out double parameter2))
            {
                MessageBox.Show(
                    "Enter a valid second parameter.");

                return;
            }


            double parameter3 =
                0;


            double parameter4 =
                0;


            // =====================================================
            // NORMAL
            // =====================================================

            if (distribution == "Normal")
            {
                if (parameter2 <= 0)
                {
                    MessageBox.Show(
                        "Standard deviation must be greater than zero.");

                    return;
                }


                Distribution =
                    DistributionType.Normal;
            }


            // =====================================================
            // LOGNORMAL
            // =====================================================

            else if (distribution == "Lognormal")
            {
                if (parameter2 <= 0)
                {
                    MessageBox.Show(
                        "Log standard deviation must be greater than zero.");

                    return;
                }


                Distribution =
                    DistributionType.Lognormal;
            }


            // =====================================================
            // UNIFORM
            // =====================================================

            else if (distribution == "Uniform")
            {
                if (parameter1 >= parameter2)
                {
                    MessageBox.Show(
                        "Minimum must be less than Maximum.");

                    return;
                }


                Distribution =
                    DistributionType.Uniform;
            }


            // =====================================================
            // BETA
            // =====================================================

            else if (distribution == "Beta")
            {
                if (parameter1 >= parameter2)
                {
                    MessageBox.Show(
                        "Minimum must be less than Maximum.");

                    return;
                }


                if (!double.TryParse(
                        txtParameter3.Text,
                        out parameter3))
                {
                    MessageBox.Show(
                        "Enter a valid Alpha value.");

                    return;
                }


                if (!double.TryParse(
                        txtParameter4.Text,
                        out parameter4))
                {
                    MessageBox.Show(
                        "Enter a valid Beta value.");

                    return;
                }


                if (parameter3 <= 0)
                {
                    MessageBox.Show(
                        "Alpha must be greater than zero.");

                    return;
                }


                if (parameter4 <= 0)
                {
                    MessageBox.Show(
                        "Beta must be greater than zero.");

                    return;
                }


                Distribution =
                    DistributionType.Beta;
            }


            // =====================================================
            // PERT / TRIANGULAR
            // =====================================================

            else
            {
                if (!double.TryParse(
                        txtParameter3.Text,
                        out parameter3))
                {
                    MessageBox.Show(
                        "Enter a valid Maximum value.");

                    return;
                }


                if (parameter1 >= parameter3)
                {
                    MessageBox.Show(
                        "Minimum must be less than Maximum.");

                    return;
                }


                if (
                    parameter2 < parameter1
                    ||
                    parameter2 > parameter3)
                {
                    MessageBox.Show(
                        "Most Likely must be between Minimum and Maximum.");

                    return;
                }


                Distribution =
                    distribution == "PERT"
                        ? DistributionType.Pert
                        : DistributionType.Triangular;
            }


            // =====================================================
            // SAVE OUTPUT VALUES
            // =====================================================

            AssumptionName =
                assumptionName;


            Parameter1 =
                parameter1;


            Parameter2 =
                parameter2;


            Parameter3 =
                parameter3;


            Parameter4 =
                parameter4;


            DialogResult =
                DialogResult.OK;


            Close();
        }
    }
}