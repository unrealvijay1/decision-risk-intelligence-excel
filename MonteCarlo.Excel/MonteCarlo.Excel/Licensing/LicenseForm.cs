using System;
using System.Drawing;
using System.Windows.Forms;

namespace MonteCarlo.Excel.Licensing
{
    public class LicenseForm : Form
    {
        private readonly Label lblLicenseValue;
        private readonly Label lblStatusValue;
        private readonly Label lblLicensedToValue;
        private readonly Label lblExpiryValue;
        private readonly Label lblMessage;

        private readonly Button btnActivate;
        private readonly Button btnDeactivate;


        public LicenseForm()
        {
            Text =
                "Monte Carlo - License";

            Width =
                560;

            Height =
                390;

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
                        "Monte Carlo for Excel",

                    Left =
                        30,

                    Top =
                        25,

                    Width =
                        470,

                    Height =
                        40,

                    Font =
                        new Font(
                            "Segoe UI",
                            16,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblTitle);


            // =====================================================
            // LICENSE TYPE
            // =====================================================

            AddCaption(
                "License:",
                90);


            lblLicenseValue =
                AddValueLabel(
                    90);


            // =====================================================
            // STATUS
            // =====================================================

            AddCaption(
                "Status:",
                125);


            lblStatusValue =
                AddValueLabel(
                    125);


            // =====================================================
            // LICENSED TO
            // =====================================================

            AddCaption(
                "Licensed To:",
                160);


            lblLicensedToValue =
                AddValueLabel(
                    160);


            // =====================================================
            // EXPIRY
            // =====================================================

            AddCaption(
                "Expiry:",
                195);


            lblExpiryValue =
                AddValueLabel(
                    195);


            // =====================================================
            // STATUS MESSAGE
            // =====================================================

            lblMessage =
                new Label
                {
                    Left =
                        30,

                    Top =
                        245,

                    Width =
                        470,

                    Height =
                        40,

                    AutoSize =
                        false
                };


            Controls.Add(
                lblMessage);


            // =====================================================
            // ACTIVATE
            // =====================================================

            btnActivate =
                new Button
                {
                    Text =
                        "Activate License",

                    Left =
                        175,

                    Top =
                        300,

                    Width =
                        125,

                    Height =
                        34,

                    AutoSize =
                        false
                };


            btnActivate.Click +=
                BtnActivate_Click;


            Controls.Add(
                btnActivate);


            // =====================================================
            // DEACTIVATE
            // =====================================================

            btnDeactivate =
                new Button
                {
                    Text =
                        "Deactivate",

                    Left =
                        310,

                    Top =
                        300,

                    Width =
                        100,

                    Height =
                        34,

                    AutoSize =
                        false
                };


            btnDeactivate.Click +=
                BtnDeactivate_Click;


            Controls.Add(
                btnDeactivate);


            // =====================================================
            // CLOSE
            // =====================================================

            Button btnClose =
                new Button
                {
                    Text =
                        "Close",

                    Left =
                        420,

                    Top =
                        300,

                    Width =
                        80,

                    Height =
                        34,

                    AutoSize =
                        false,

                    DialogResult =
                        DialogResult.OK
                };


            Controls.Add(
                btnClose);


            AcceptButton =
                btnClose;


            CancelButton =
                btnClose;


            // =====================================================
            // INITIAL DISPLAY
            // =====================================================

            RefreshLicenseDisplay();
        }


        // =========================================================
        // ADD CAPTION
        // =========================================================

        private void AddCaption(
            string caption,
            int top)
        {
            Label label =
                new Label
                {
                    Text =
                        caption,

                    Left =
                        30,

                    Top =
                        top,

                    Width =
                        120,

                    Height =
                        24,

                    Font =
                        new Font(
                            "Segoe UI",
                            9,
                            FontStyle.Bold)
                };


            Controls.Add(
                label);
        }


        // =========================================================
        // ADD VALUE LABEL
        // =========================================================

        private Label AddValueLabel(
            int top)
        {
            Label label =
                new Label
                {
                    Left =
                        155,

                    Top =
                        top,

                    Width =
                        345,

                    Height =
                        24,

                    AutoSize =
                        false
                };


            Controls.Add(
                label);


            return
                label;
        }


        // =========================================================
        // REFRESH LICENSE DISPLAY
        // =========================================================

        private void RefreshLicenseDisplay()
        {
            LicenseInfo license =
                LicenseService.GetCurrentLicense();


            lblLicenseValue.Text =
                license.Type.ToString();


            lblStatusValue.Text =
                license.Status.ToString();


            lblLicensedToValue.Text =
                string.IsNullOrWhiteSpace(
                    license.LicensedTo)
                    ? "-"
                    : license.LicensedTo;


            if (license.ExpiryDate.HasValue)
            {
                string expiryText =
                    license.ExpiryDate.Value
                        .ToString(
                            "dd-MMM-yyyy");


                if (
                    license.Type ==
                    LicenseType.Trial
                    &&
                    license.Status ==
                    LicenseStatus.Active)
                {
                    expiryText +=
                        $" ({license.DaysRemaining} days remaining)";
                }


                lblExpiryValue.Text =
                    expiryText;
            }
            else
            {
                lblExpiryValue.Text =
                    "No expiry";
            }


            lblMessage.Text =
                GetStatusMessage(
                    license);


            // =====================================================
            // BUTTON STATES
            // =====================================================

            if (
                license.Type ==
                LicenseType.Development)
            {
                btnActivate.Enabled =
                    false;


                btnDeactivate.Enabled =
                    false;


                return;
            }


            if (
                license.Type ==
                LicenseType.Professional
                ||
                license.Type ==
                LicenseType.Enterprise)
            {
                btnActivate.Enabled =
                    true;


                btnDeactivate.Enabled =
                    true;


                return;
            }


            // Trial / expired / invalid
            btnActivate.Enabled =
                true;


            btnDeactivate.Enabled =
                LicenseStorage.HasSavedLicense();
        }


        // =========================================================
        // ACTIVATE
        // =========================================================

        private void BtnActivate_Click(
            object? sender,
            EventArgs e)
        {
            try
            {
                using ActivationForm activationForm =
                    new ActivationForm();


                if (
                    activationForm.ShowDialog(this)
                    != DialogResult.OK)
                {
                    return;
                }


                RefreshLicenseDisplay();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Monte Carlo License",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }


        // =========================================================
        // DEACTIVATE
        // =========================================================

        private void BtnDeactivate_Click(
            object? sender,
            EventArgs e)
        {
            try
            {
                DialogResult confirmation =
                    MessageBox.Show(
                        "Deactivate the current license on this computer?",
                        "Monte Carlo License",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);


                if (
                    confirmation !=
                    DialogResult.Yes)
                {
                    return;
                }


                LicenseService.Deactivate();


                RefreshLicenseDisplay();


                MessageBox.Show(
                    "The local license has been deactivated.",
                    "Monte Carlo License",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Monte Carlo License",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }


        // =========================================================
        // STATUS MESSAGE
        // =========================================================

        private static string GetStatusMessage(
            LicenseInfo license)
        {
            if (
                license.Type ==
                LicenseType.Development)
            {
                return
                    "Development mode - all features are enabled.";
            }


            if (
                license.Type ==
                LicenseType.Trial
                &&
                license.Status ==
                LicenseStatus.Active)
            {
                return
                    $"Trial version - " +
                    $"{license.DaysRemaining} days remaining.";
            }


            if (
                license.Type ==
                LicenseType.Professional
                &&
                license.Status ==
                LicenseStatus.Active)
            {
                return
                    "Professional license is active.";
            }


            if (
                license.Type ==
                LicenseType.Enterprise
                &&
                license.Status ==
                LicenseStatus.Active)
            {
                return
                    "Enterprise license is active.";
            }


            if (
                license.Status ==
                LicenseStatus.Expired)
            {
                return
                    "Your trial or license has expired. " +
                    "Activate a valid license to continue using simulations.";
            }


            if (
                license.Status ==
                LicenseStatus.Invalid)
            {
                return
                    "The current license is not valid.";
            }


            return
                "License status is unavailable.";
        }
    }
}