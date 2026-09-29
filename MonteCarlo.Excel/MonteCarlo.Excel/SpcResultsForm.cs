using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public sealed class SpcResultsForm : Form
{
    public SpcResultsForm(SpcResult result, Action export)
    {
        Text = "SPC Results — " + result.Configuration.ProcessName;
        Font = new Font("Segoe UI", 9); AutoScaleMode = AutoScaleMode.Dpi;
        Size = new Size(1160, 900); MinimumSize = new Size(900, 740); StartPosition = FormStartPosition.CenterParent;
        var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 5 };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 40)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 20)); root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        Controls.Add(root);
        var b = result.Baseline; var c = result.Configuration;
        string target = c.Target is { } t ? $"{t.Value:G6} ({(t.Direction == TargetDirection.AtOrBelow ? "lower" : "higher")} is better)" : "Not configured";
        string Short(string text) => text.Length <= 100 ? text : text[..97] + "…";
        var summary = new Label { Name = "SpcSummary", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0, 0, 0, 10),
            BackColor = result.BaselineHasSignals ? Color.LightYellow : SystemColors.Control, Text =
            $"{Short(c.ProcessName)} — {result.Status}" + (result.BaselineHasSignals ? "  |  BASELINE REQUIRES INVESTIGATION" : "") + Environment.NewLine +
            $"Observations: {result.Observations.Count}    Baseline: {b.End - b.Start + 1} observations, #{b.Start}–#{b.End} ({Short(result.Labels[b.Start - 1])} → {Short(result.Labels[b.End - 1])})" + Environment.NewLine +
            $"Mean: {b.Mean:G6}    Average moving range: {b.AverageMovingRange:G6}    Individuals LCL: {b.IndividualsLower:G6}    UCL: {b.IndividualsUpper:G6}" + Environment.NewLine +
            $"Latest: {result.Observations[^1]:G6} {c.MeasurementUnit}    Business target: {target}    Signals: {result.Signals.Count}" + Environment.NewLine +
            "Fixed baseline limits: X̄ ± 2.66 × MR̄; moving range UCL = 3.268 × MR̄, LCL = 0." };
        root.Controls.Add(summary, 0, 0);
        root.SizeChanged += (_, _) => summary.MaximumSize = new Size(Math.Max(200, root.ClientSize.Width - 32), 0);
        root.Controls.Add(new SpcChartPanel(result, SpcChart.Individuals), 0, 1);
        root.Controls.Add(new SpcChartPanel(result, SpcChart.MovingRange), 0, 2);
        var tabs = new TabControl { Dock = DockStyle.Fill };
        var insights = new TabPage("Process Insights");
        insights.Controls.Add(new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Text = string.Join(Environment.NewLine, result.Insights()) });
        var signals = new TabPage("Signal Details");
        var grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
            DataSource = result.Signals.Select(s => new { s.RuleId, s.RuleName, s.Chart, s.Direction, s.StartIndex, s.EndIndex,
                s.QualifyingObservations, Threshold = s.ReferenceValue, WindowValues = s.Values, s.Description }).ToArray() };
        grid.DataBindingComplete += (_, _) =>
        {
            foreach (var (name, caption) in new[] { ("RuleId", "Rule"), ("RuleName", "Rule name"), ("StartIndex", "Window start"),
                ("EndIndex", "Window end"), ("QualifyingObservations", "Qualifying observations"), ("WindowValues", "Actual values (window order)") })
                if (grid.Columns[name] is { } column) column.HeaderText = caption;
            grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        };
        signals.Controls.Add(grid); tabs.TabPages.Add(insights); tabs.TabPages.Add(signals); root.Controls.Add(tabs, 0, 3);
        var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        var close = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.Cancel };
        var save = new Button { Text = "Export to Worksheet", AutoSize = true };
        save.Click += (_, _) => { try { export(); } catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); MessageBox.Show(this, "Export could not be completed. Check that the source workbook is open, writable and unprotected. A partial new report may remain.", "SPC Export"); } };
        footer.Controls.Add(close); footer.Controls.Add(save); root.Controls.Add(footer, 0, 4); CancelButton = close;
    }
}
