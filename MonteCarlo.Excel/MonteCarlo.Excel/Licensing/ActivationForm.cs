using System;
using System.Drawing;
using System.Windows.Forms;

namespace MonteCarlo.Excel.Licensing
{
    public class ActivationForm : Form
    {
        private readonly TextBox txtLicenseKey;
        private readonly Label lblStatus;


        public ActivationForm()
        {
            Text =
                "Activate Monte Carlo";

            Width =
                500;

            Height =
                300;

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
                        "Activate Monte Carlo for Excel",

                    Left =
                        25,

                    Top =
                        20,

                    Width =
                        420,

                    Height =
                        35,

                    Font =
                        new Font(
                            "Segoe UI",
                            15,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblTitle);


            // =====================================================
            // INSTRUCTIONS
            // =====================================================

            Label lblInstructions =
                new Label
                {
                    Text =
                        "Enter the license key provided when you purchased Monte Carlo.",

                    Left =
                        25,

                    Top =
                        65,

                    Width =
                        430,

                    Height =
                        40
                };


            Controls.Add(
                lblInstructions);


            // =====================================================
            // LICENSE KEY LABEL
            // =====================================================

            Label lblLicenseKey =
                new Label
                {
                    Text =
                        "License Key",

                    Left =
                        25,

                    Top =
                        120,

                    Width =
                        100,

                    Height =
                        25,

                    Font =
                        new Font(
                            "Segoe UI",
                            9,
                            FontStyle.Bold)
                };


            Controls.Add(
                lblLicenseKey);


            // =====================================================
            // LICENSE KEY TEXT BOX
            // =====================================================

            txtLicenseKey =
                new TextBox
                {
                    Left =
                        130,

                    Top =
                        115,

                    Width =
                        320,

                    CharacterCasing =
                        CharacterCasing.Upper
                };


            Controls.Add(
                txtLicenseKey);


            // =====================================================
            // STATUS
            // =====================================================

            lblStatus =
                new Label
                {
                    Left =
                        25,

                    Top =
                        160,

                    Width =
                        425,

                    Height =
                        35
                };


            Controls.Add(
                lblStatus);


            // =====================================================
            // ACTIVATE
            // =====================================================

            Button btnActivate =
                new Button
                {
                    Text =
                        "Activate",

                    Left =
                        270,

                    Top =
                        205,

                    Width =
                        85,

                    Height =
                        32
                };


            btnActivate.Click +=
                BtnActivate_Click;


            Controls.Add(
                btnActivate);


            // =====================================================
            // CANCEL
            // =====================================================

            Button btnCancel =
                new Button
                {
                    Text =
                        "Cancel",

                    Left =
                        365,

                    Top =
                        205,

                    Width =
                        85,

                    Height =
                        32,

                    DialogResult =
                        DialogResult.Cancel
                };


            Controls.Add(
                btnCancel);


            AcceptButton =
                btnActivate;


            CancelButton =
                btnCancel;


            // =====================================================
            // LOAD EXISTING KEY IF PRESENT
            // =====================================================

            string? existingKey =
                LicenseStorage.LoadLicenseKey();


            if (!string.IsNullOrWhiteSpace(
                    existingKey))
            {
                txtLicenseKey.Text =
                    existingKey;
            }
        }


        // =========================================================
        // ACTIVATE
        // =========================================================

        private void BtnActivate_Click(
            object? sender,
            EventArgs e)
        {
            string licenseKey =
                txtLicenseKey.Text.Trim();


            if (string.IsNullOrWhiteSpace(
                    licenseKey))
            {
                MessageBox.Show(
                    "Enter a license key.",
                    "Monte Carlo Activation",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }


            try
            {
                lblStatus.Text =
                    "Validating license...";


                Application.DoEvents();


                LicenseInfo? activatedLicense =
                    LicenseService.Activate(
                        licenseKey);


                if (
                    activatedLicense == null
                    ||
                    !activatedLicense.IsValid)
                {
                    lblStatus.Text =
                        "License key is not valid.";


                    MessageBox.Show(
                        "The license key could not be activated.",
                        "Monte Carlo Activation",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }


                LicenseStorage.SaveLicenseKey(
                    licenseKey);


                lblStatus.Text =
                    "License activated successfully.";


                MessageBox.Show(
                    $"License activated successfully.\n\n" +
                    $"Type: {activatedLicense.Type}\n" +
                    $"Licensed To: {activatedLicense.LicensedTo}",
                    "Monte Carlo Activation",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);


                DialogResult =
                    DialogResult.OK;


                Close();
            }
            catch (Exception ex)
            {
                lblStatus.Text =
                    "Activation failed.";


                MessageBox.Show(
                    ex.Message,
                    "Monte Carlo Activation",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}