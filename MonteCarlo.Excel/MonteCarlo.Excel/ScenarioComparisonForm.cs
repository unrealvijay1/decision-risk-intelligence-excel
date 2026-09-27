using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public sealed class ScenarioComparisonForm : Form
{
    public ScenarioComparisonForm(ScenarioComparisonResult result)
    {
        Text = "Scenario Comparison"; AutoScaleMode = AutoScaleMode.Dpi; ClientSize = new(1100, 560);
        MinimumSize = new(760, 400); StartPosition = FormStartPosition.CenterParent;
        bool valueMode = result.Mode == ScenarioAnalysisMode.ProbabilityToValue;
        var root = ScenarioUi.Table(1, 4); root.Padding = new Padding(12); Controls.Add(root);
        root.RowStyles.Add(new(SizeType.AutoSize)); root.RowStyles.Add(new(SizeType.Absolute, 95));
        root.RowStyles.Add(new(SizeType.Percent, 100)); root.RowStyles.Add(new(SizeType.AutoSize));
        root.Controls.Add(new Label { AutoSize = true, Name = "ComparisonQuestion", Text =
            $"{result.Trials:N0} trials per scenario · Common seed: {result.Seed}\nBaseline: {result.BaselineName}\nQuestion: " +
            (valueMode ? "Value at configured probabilities per forecast" : "Probability of meeting configured targets") }, 0, 0);
        var context = new TextBox { Name = "ForecastContext", Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
            Text = string.Join(Environment.NewLine, result.Rows.Where(r => r.IsBaseline).Select(r =>
                $"{r.Forecast}: " + (valueMode ? $"Value at {r.Probability ?? result.Probability:G}% probability, " : $"Target {r.Target:G} ") +
                (r.Direction == TargetDirection.AtOrBelow ? "at or below (≤)" : "at or above (≥)"))) };
        root.Controls.Add(context, 0, 1);
        var grid = ScenarioUi.Grid("ScenarioComparison"); grid.ReadOnly = true;
        root.Controls.Add(grid, 0, 2);
        string[] columns = { "Scenario", "Forecast", "Mean", "Δ Mean", valueMode ? "Probability Value" : "Probability", valueMode ? "Δ Value" : "Δ Probability (pp)" };
        foreach (string column in columns) grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = column,
            FillWeight = column is "Scenario" or "Forecast" ? 150 : 100, MinimumWidth = 90, SortMode = DataGridViewColumnSortMode.NotSortable });
        string Delta(double? value, string suffix = "") => value.HasValue ? value.Value.ToString("+#,##0.00;-#,##0.00;0.00") + suffix : "—";
        foreach (var row in result.Rows)
        {
            int index = grid.Rows.Add(row.Scenario + (row.IsBaseline ? " ★ Baseline" : ""), row.Forecast, row.Mean.ToString("N2"), Delta(row.MeanDelta),
                valueMode ? row.DecisionValue.ToString("N2") : row.DecisionValue.ToString("P1"), Delta(row.DecisionDelta, valueMode ? "" : " pp"));
            if (row.IsBaseline) grid.Rows[index].DefaultCellStyle.BackColor = Color.FromArgb(235, 242, 250);
        }
        var footer = ScenarioUi.Flow(); var close = ScenarioUi.Button("Close"); close.DialogResult = DialogResult.OK;
        footer.Controls.Add(close); root.Controls.Add(footer, 0, 3); AcceptButton = CancelButton = close;
    }
}
