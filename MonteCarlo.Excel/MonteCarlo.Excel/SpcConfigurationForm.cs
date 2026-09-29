using System.Globalization;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public sealed record SpcDraft(SpcConfiguration Configuration, bool ReplaceBaseline);

public sealed class SpcConfigurationForm : Form
{
    private readonly TextBox process = new() { Dock = DockStyle.Fill };
    private readonly TextBox unit = new() { Dock = DockStyle.Fill };
    private readonly TextBox target = new() { Dock = DockStyle.Fill };
    private readonly ComboBox mode = new() { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox direction = new() { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly NumericUpDown start = new() { Minimum = 1, Maximum = SpcDataReader.MaximumObservations, Value = 1, Dock = DockStyle.Fill };
    private readonly NumericUpDown end = new() { Minimum = 20, Maximum = SpcDataReader.MaximumObservations, Value = 20, Dock = DockStyle.Fill };
    private readonly CheckBox beyond = new() { Text = "Beyond control limits (3σ)", Checked = true, AutoSize = true };
    private readonly CheckBox twoOfThree = new() { Text = "Two of three beyond 2σ", Checked = true, AutoSize = true };
    private readonly CheckBox fourOfFive = new() { Text = "Four of five beyond 1σ", Checked = true, AutoSize = true };
    private readonly CheckBox run = new() { Text = "Eight-point run", Checked = true, AutoSize = true };
    private readonly CheckBox trend = new() { Text = "Six-point trend", Checked = true, AutoSize = true };
    private readonly Label baselineInfo = new() { AutoSize = true, MaximumSize = new Size(650, 0) };
    private readonly Label validation = new() { AutoSize = true, ForeColor = Color.Firebrick, MaximumSize = new Size(650, 0) };
    public SpcConfigurationForm(SpcSavedAnalysis? saved, string dataDescription, string labelDescription,
        Func<bool, string?> pickRange, Action clearLabels, Action<SpcDraft> analyze)
    {
        Text = "SPC Analysis — Process Analysis"; Font = new Font("Segoe UI", 10);
        AutoScaleMode = AutoScaleMode.Dpi; Size = new Size(800, 790); MinimumSize = new Size(740, 650); StartPosition = FormStartPosition.CenterParent;
        var outer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(14) };
        outer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        outer.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); outer.RowStyles.Add(new RowStyle(SizeType.Absolute, 82)); Controls.Add(outer);
        var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true }; outer.Controls.Add(scroll, 0, 0);
        var table = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, Padding = new Padding(0, 0, 18, 0) };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170)); table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); scroll.Controls.Add(table);
        void Row(string caption, Control control)
        {
            int row = table.RowCount++; table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(new Label { Text = caption, AutoSize = true, Margin = new Padding(3, 8, 3, 10) }, 0, row);
            control.Margin = new Padding(3, 5, 3, 8); table.Controls.Add(control, 1, row);
        }
        Control RangeRow(bool labels, string description)
        {
            var panel = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = true };
            var text = new Label { Text = description, AutoSize = true, MaximumSize = new Size(410, 0) };
            var select = new Button { Text = "Select in Excel…", AutoSize = true };
            select.Click += (_, _) =>
            {
                try { Hide(); string? selected = pickRange(labels); if (selected != null) text.Text = selected; }
                catch (ArgumentException ex) { validation.Text = ex.Message; }
                catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); validation.Text = "Could not select the range. Check the workbook and try again."; }
                finally { Show(); Activate(); }
            };
            panel.Controls.Add(select);
            if (labels)
            {
                var clear = new Button { Text = "Clear", AutoSize = true }; clear.Click += (_, _) => { clearLabels(); text.Text = "Observation numbers"; }; panel.Controls.Add(clear);
            }
            panel.Controls.Add(text); return panel;
        }
        Row("Data range", RangeRow(false, dataDescription)); Row("Date/sequence range", RangeRow(true, labelDescription));
        Row("Process name", process); Row("Measurement unit", unit);
        mode.Items.AddRange(["All observations", "First N observations", "Selected contiguous period"]);
        Row("Baseline", mode); Row("First observation", start); Row("Last observation / N", end);
        Row("Baseline period", baselineInfo);
        Row("Baseline guidance", new Label { AutoSize = true, MaximumSize = new Size(470, 0), Text = "At least 20 observations are required. This is a practical minimum, not a guarantee of reliable limits. Analyze reuses established limits; replace the baseline deliberately to recalculate." });
        Row("Business target", target); direction.Items.AddRange(["Lower is better (≤)", "Higher is better (≥)"]); Row("Target direction", direction);
        Row("Target guidance", new Label { AutoSize = true, MaximumSize = new Size(470, 0), Text = "Optional. A business target is separate from control limits. Leave blank for no target." });
        var rules = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };
        rules.Controls.AddRange([beyond, run, twoOfThree, fourOfFive]); Row("Wheeler's Four Tests", rules);
        Row("Supplementary Rules", trend);
        Row("Validation", validation);
        var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = true };
        var close = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.Cancel };
        var go = new Button { Text = "Analyze", AutoSize = true, Enabled = saved != null };
        var establish = new Button { Text = "Establish / replace baseline", AutoSize = true };
        footer.Controls.AddRange([close, go, establish]); outer.Controls.Add(footer, 0, 1); CancelButton = close; AcceptButton = go;
        void Submit(bool replace)
        {
            try
            {
                validation.Text = ""; analyze(new(ReadConfiguration(), replace)); go.Enabled = true;
                baselineInfo.Text = "Saved baseline established. Analyze keeps its limits fixed. Save the workbook to retain it on disk.";
            }
            catch (ArgumentException ex) { validation.Text = ex.Message; }
            catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); validation.Text = "SPC analysis could not complete. Check source references and workbook write protection, then try again."; }
        }
        go.Click += (_, _) => Submit(false); establish.Click += (_, _) => Submit(true);
        var config = saved?.Configuration ?? new(); process.Text = config.ProcessName; unit.Text = config.MeasurementUnit;
        mode.SelectedIndex = (int)config.BaselineMode; start.Value = Math.Clamp(config.BaselineStart, 1, SpcDataReader.MaximumObservations);
        end.Value = Math.Clamp(config.BaselineEnd, 20, SpcDataReader.MaximumObservations);
        target.Text = config.Target?.Value.ToString("R", CultureInfo.CurrentCulture) ?? ""; direction.SelectedIndex = config.Target?.Direction == TargetDirection.AtOrAbove ? 1 : 0;
        beyond.Checked = config.Rules.BeyondLimits; run.Checked = config.Rules.SustainedRun; trend.Checked = config.Rules.SustainedTrend;
        twoOfThree.Checked = config.Rules.TwoOfThree; fourOfFive.Checked = config.Rules.FourOfFive;
        void BaselineDescription()
        {
            start.Enabled = mode.SelectedIndex == 2; end.Enabled = mode.SelectedIndex != 0;
            baselineInfo.Text = saved is null ? "No baseline established. " : $"Saved limits: observations #{saved.Baseline.Start}–#{saved.Baseline.End}. ";
            baselineInfo.Text += mode.SelectedIndex == 0 ? "Replacement uses all selected observations." : $"Replacement uses #{(mode.SelectedIndex == 1 ? 1 : start.Value)}–#{end.Value}, inclusive.";
        }
        mode.SelectedIndexChanged += (_, _) => BaselineDescription(); start.ValueChanged += (_, _) => BaselineDescription(); end.ValueChanged += (_, _) => BaselineDescription(); BaselineDescription();
    }
    public SpcConfiguration ReadConfiguration()
    {
        if (string.IsNullOrWhiteSpace(process.Text)) throw new ArgumentException("Enter a process name.");
        SpcTarget? businessTarget = null;
        if (!string.IsNullOrWhiteSpace(target.Text))
        {
            if (!double.TryParse(target.Text, out double value) || !double.IsFinite(value)) throw new ArgumentException("Enter a finite numerical target or leave it blank.");
            businessTarget = new(value, direction.SelectedIndex == 0 ? TargetDirection.AtOrBelow : TargetDirection.AtOrAbove);
        }
        return new() { ProcessName = process.Text.Trim(), MeasurementUnit = unit.Text.Trim(), BaselineMode = (SpcBaselineMode)mode.SelectedIndex,
            BaselineStart = mode.SelectedIndex == 2 ? (int)start.Value : 1, BaselineEnd = (int)end.Value,
            Target = businessTarget, Rules = new(beyond.Checked, run.Checked, trend.Checked, twoOfThree.Checked, fourOfFive.Checked) };
    }
}
