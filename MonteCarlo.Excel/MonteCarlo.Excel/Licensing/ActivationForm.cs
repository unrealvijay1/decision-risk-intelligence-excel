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
                650;

            Height =
                390;

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
                        "Activate Monte Carlo for Excel",

                    Left =
                        30,

                    Top =
                        25,

                    Width =
                        560,

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
            // INSTRUCTIONS
            // =====================================================

            Label lblInstructions =
                new Label
                {
                    Text =
                        "Paste the complete license code provided " +
                        "when you purchased Monte Carlo.",

                    Left =
                        30,

                    Top =
                        75,

                    Width =
                        560,

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
                        "License Code",

                    Left =
                        30,

                    Top =
                        130,

                    Width =
                        110,

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
            // LICENSE CODE TEXT BOX
            //
            // IMPORTANT:
            // Do NOT uppercase this text.
            //
            // RSA signed license codes are case-sensitive.
            // =====================================================

            txtLicenseKey =
                new TextBox
                {
                    Left =
                        145,

                    Top =
                        125,

                    Width =
                        440,

                    Height =
                        75,

                    Multiline =
                        true,

                    WordWrap =
                        false,

                    ScrollBars =
                        ScrollBars.Horizontal,

                    CharacterCasing =
                        CharacterCasing.Normal
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
                        30,

                    Top =
                        220,

                    Width =
                        555,

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
                        395,

                    Top =
                        285,

                    Width =
                        90,

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
                        495,

                    Top =
                        285,

                    Width =
                        90,

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
            // LOAD EXISTING LICENSE IF PRESENT
            // =====================================================

            string? existingKey =
                LicenseStorage.LoadLicenseKey();


            if (!string.IsNullOrWhiteSpace(
                    existingKey))
            {
                txtLicenseKey.Text =
                    existingKey;


                txtLicenseKey.SelectionStart =
                    0;


                txtLicenseKey.SelectionLength =
                    0;
            }
        }


        // =========================================================
        // ACTIVATE
        // =========================================================

        private void BtnActivate_Click(
            object? sender,
            EventArgs e)
        {
            string licenseCode =
                txtLicenseKey.Text.Trim();


            if (string.IsNullOrWhiteSpace(
                    licenseCode))
            {
                MessageBox.Show(
                    "Paste the complete license code.",
                    "Monte Carlo Activation",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }


            // =====================================================
            // BASIC FORMAT CHECK
            // =====================================================

            if (!licenseCode.StartsWith(
                    "MC1.",
                    StringComparison.Ordinal))
            {
                MessageBox.Show(
                    "This does not appear to be a valid Monte Carlo " +
                    "license code.\n\n" +
                    "The license code should begin with MC1.",
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
                        licenseCode);


                if (activatedLicense == null)
                {
                    lblStatus.Text =
                        "License code is not valid.";


                    MessageBox.Show(
                        "The license could not be validated.\n\n" +
                        "Make sure the entire license code was copied " +
                        "without modifying any characters.",
                        "Monte Carlo Activation",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }


                // =================================================
                // EXPIRED LICENSE
                // =================================================

                if (!activatedLicense.IsValid)
                {
                    lblStatus.Text =
                        "License has expired.";


                    MessageBox.Show(
                        $"This license has expired.\n\n" +
                        $"Licensed To: " +
                        $"{activatedLicense.LicensedTo}\n" +
                        $"Expiry: " +
                        $"{activatedLicense.ExpiryDate:dd-MMM-yyyy}",
                        "Monte Carlo Activation",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }


                // =================================================
                // SUCCESS
                // =================================================

                lblStatus.Text =
                    "License activated successfully.";


                MessageBox.Show(
                    $"License activated successfully.\n\n" +
                    $"Type: {activatedLicense.Type}\n" +
                    $"Licensed To: " +
                    $"{activatedLicense.LicensedTo}\n" +
                    $"Expiry: " +
                    $"{activatedLicense.ExpiryDate:dd-MMM-yyyy}",
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