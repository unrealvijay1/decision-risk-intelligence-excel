using System;
using System.Windows.Forms;

namespace MonteCarlo.Excel
{
    public class SimulationSettingsForm : Form
    {
        private readonly NumericUpDown numTrials;

        public int Trials { get; private set; }

        public SimulationSettingsForm()
        {
            Text = "Simulation Settings";

            Width = 360;
            Height = 220;

            StartPosition =
                FormStartPosition.CenterScreen;

            FormBorderStyle =
                FormBorderStyle.FixedDialog;

            MaximizeBox = false;
            MinimizeBox = false;


            Label lblTrials = new()
            {
                Text = "Number of Trials",
                Left = 25,
                Top = 35,
                Width = 120
            };


            numTrials = new NumericUpDown
            {
                Left = 160,
                Top = 30,
                Width = 140,

                Minimum = 100,
                Maximum = 1000000,

                Value = 10000,

                Increment = 1000,

                ThousandsSeparator = true
            };


            Label lblHint = new()
            {
                Text =
                    "Recommended: 10,000 trials",
                Left = 160,
                Top = 65,
                Width = 160
            };


            Button btnRun = new()
            {
                Text = "Run",
                Left = 140,
                Top = 115,
                Width = 75
            };


            Button btnCancel = new()
            {
                Text = "Cancel",
                Left = 225,
                Top = 115,
                Width = 75
            };


            btnRun.Click +=
                (_, _) =>
                {
                    Trials =
                        (int)numTrials.Value;

                    DialogResult =
                        DialogResult.OK;

                    Close();
                };


            btnCancel.Click +=
                (_, _) =>
                {
                    DialogResult =
                        DialogResult.Cancel;

                    Close();
                };


            Controls.Add(
                lblTrials);

            Controls.Add(
                numTrials);

            Controls.Add(
                lblHint);

            Controls.Add(
                btnRun);

            Controls.Add(
                btnCancel);


            AcceptButton =
                btnRun;

            CancelButton =
                btnCancel;
        }
    }
}