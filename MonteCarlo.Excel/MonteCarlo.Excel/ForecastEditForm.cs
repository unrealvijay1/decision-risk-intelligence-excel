using System;
using System.Drawing;
using System.Windows.Forms;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public sealed class ForecastEditForm : Form
{
    public ForecastEditDraft Draft { get; }

    public ForecastEditForm(ForecastDefinition forecast)
    {
        Draft = new ForecastEditDraft(forecast);
        Text = "Edit Forecast";
        ClientSize = new Size(510, 330);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = MinimizeBox = false;
        AutoScaleMode = AutoScaleMode.Font;

        var location = new Label { Text = $"Cell: {forecast.SheetName}!{forecast.CellAddress}", Left = 20, Top = 15, Width = 470 };
        var name = new TextBox { Text = Draft.Name, Left = 165, Top = 48, Width = 320 };
        var mode = new ComboBox { Left = 165, Top = 88, Width = 320, DropDownStyle = ComboBoxStyle.DropDownList };
        mode.Items.AddRange(new object[] { "Default (P80 from simulation)", "Specified target", "No target" });
        mode.SelectedIndex = (int)Draft.Mode;
        var target = new TextBox { Text = Draft.TargetText, Left = 165, Top = 128, Width = 320 };
        var direction = new ComboBox { Left = 165, Top = 168, Width = 320, DropDownStyle = ComboBoxStyle.DropDownList };
        direction.Items.AddRange(new object[] { "Unspecified", "At or below target", "At or above target" });
        direction.SelectedIndex = Draft.Direction switch { TargetDirection.AtOrBelow => 1, TargetDirection.AtOrAbove => 2, _ => 0 };
        var note = new Label { Left = 20, Top = 210, Width = 465, Height = 55,
            Text = Draft.RequestedConfidence.HasValue
                ? $"Saved confidence: {Draft.RequestedConfidence.Value:G}% (retained unless target or direction changes)."
                : "Default uses P80 when results are calculated. No target explicitly clears the target." };
        void UpdateFields()
        {
            target.Enabled = mode.SelectedIndex == (int)ForecastTargetMode.Manual;
            direction.Enabled = mode.SelectedIndex != (int)ForecastTargetMode.DefaultP80;
        }
        mode.SelectedIndexChanged += (_, _) => UpdateFields();
        UpdateFields();
        var save = new Button { Text = "Save", Left = 315, Top = 280, Width = 80 };
        var cancel = new Button { Text = "Cancel", Left = 405, Top = 280, Width = 80, DialogResult = DialogResult.Cancel };
        save.Click += (_, _) =>
        {
            Draft.Name = name.Text;
            Draft.Mode = (ForecastTargetMode)mode.SelectedIndex;
            Draft.TargetText = target.Text;
            Draft.Direction = direction.SelectedIndex switch { 1 => TargetDirection.AtOrBelow, 2 => TargetDirection.AtOrAbove, _ => null };
            // Validate against a detached definition; the caller saves only after OK.
            if (!Draft.TryApply(new ForecastDefinition(), out string? error))
            {
                MessageBox.Show(this, error, "Edit Forecast", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            DialogResult = DialogResult.OK;
            Close();
        };
        Controls.AddRange(new Control[] { location, name, mode, target, direction, note, save, cancel });
        string[] labels = { "Forecast name", "Target mode", "Target", "Success direction" };
        for (int i = 0; i < labels.Length; i++)
            Controls.Add(new Label { Text = labels[i], Left = 20, Top = 51 + i * 40, Width = 140 });
        AcceptButton = save;
        CancelButton = cancel;
    }
}
