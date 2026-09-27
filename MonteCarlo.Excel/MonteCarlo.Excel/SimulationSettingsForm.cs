using System.Windows.Forms;

namespace MonteCarlo.Excel;

public sealed class SimulationSettingsForm : Form
{
    public SimulationSettings Settings { get; private set; }
    public SimulationSettingsForm(SimulationSettings settings)
    {
        Settings = settings;
        Text = "Simulation Settings"; ClientSize = new System.Drawing.Size(430, 255);
        FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        var preset = new ComboBox { Left = 175, Top = 20, Width = 220, DropDownStyle = ComboBoxStyle.DropDownList };
        foreach (int count in SimulationSettings.Presets) preset.Items.Add(count.ToString("N0"));
        preset.Items.Add("Custom");
        var trials = new TextBox { Left = 175, Top = 60, Width = 220, Text = settings.TrialCount.ToString() };
        preset.SelectedIndex = SimulationSettings.Presets.ToList().IndexOf(settings.TrialCount);
        if (preset.SelectedIndex < 0) preset.SelectedIndex = 4;
        trials.Enabled = preset.SelectedIndex == 4;
        preset.SelectedIndexChanged += (_, _) => {
            trials.Enabled = preset.SelectedIndex == 4;
            if (!trials.Enabled) trials.Text = SimulationSettings.Presets[preset.SelectedIndex].ToString();
        };
        var mode = new ComboBox { Left = 175, Top = 100, Width = 220, DropDownStyle = ComboBoxStyle.DropDownList };
        mode.Items.AddRange(new object[] { "Automatic", "Fixed" }); mode.SelectedIndex = (int)settings.SeedMode;
        var seed = new TextBox { Left = 175, Top = 140, Width = 220, Text = settings.FixedSeed?.ToString() ?? "", Enabled = mode.SelectedIndex == 1 };
        mode.SelectedIndexChanged += (_, _) => seed.Enabled = mode.SelectedIndex == 1;
        var save = new Button { Text = "Save", Left = 225, Top = 200, Width = 80 };
        var cancel = new Button { Text = "Cancel", Left = 315, Top = 200, Width = 80, DialogResult = DialogResult.Cancel };
        save.Click += (_, _) => {
            if (!SimulationSettings.TryParse(trials.Text, (SimulationSeedMode)mode.SelectedIndex, seed.Text, out var parsed, out var error))
            { MessageBox.Show(this, error, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            Settings = parsed; DialogResult = DialogResult.OK; Close();
        };
        Controls.AddRange(new Control[] { preset, trials, mode, seed, save, cancel });
        string[] labels = { "Trials preset", "Number of trials", "Seed mode", "Fixed seed" };
        for (int i = 0; i < labels.Length; i++) Controls.Add(new Label { Text = labels[i], Left = 20, Top = 24 + 40 * i, Width = 150 });
        AcceptButton = save; CancelButton = cancel;
    }
}
