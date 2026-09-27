using System.Globalization;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public sealed class ScenarioQuestionControl : UserControl
{
    private readonly IReadOnlyList<ForecastDefinition> forecasts;
    private List<ScenarioForecastQuestion> questions;
    private readonly ComboBox mode = new() { Name = "AnalysisMode", DropDownStyle = ComboBoxStyle.DropDownList, Width = 230 };
    private readonly double legacyProbability;
    private readonly DataGridView grid = ScenarioUi.Grid("ForecastQuestions");
    private ScenarioAnalysisMode currentMode;
    private bool loading;
    public event Action? Changed;

    public ScenarioQuestionControl(IReadOnlyList<ForecastDefinition> forecasts, ScenarioAnalysisDefinition analysis)
    {
        legacyProbability = analysis.Probability; this.forecasts = forecasts; questions = analysis.Questions.ToList(); currentMode = analysis.Mode;
        Dock = DockStyle.Fill;
        var root = ScenarioUi.Table(1, 3); Controls.Add(root);
        root.RowStyles.Add(new(SizeType.AutoSize)); root.RowStyles.Add(new(SizeType.Percent, 100)); root.RowStyles.Add(new(SizeType.AutoSize));
        var top = ScenarioUi.Flow(); root.Controls.Add(top, 0, 0);
        mode.Items.AddRange(new object[] { "Probability → Value", "Value → Probability" }); mode.SelectedIndex = (int)currentMode;

        top.Controls.Add(mode);
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Forecast", ReadOnly = true, FillWeight = 50 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Target value", FillWeight = 25 });
        var directions = new DataGridViewComboBoxColumn { HeaderText = "Direction", FillWeight = 25 };
        directions.Items.AddRange("Select direction", "At or Below (≤)", "At or Above (≥)"); grid.Columns.Add(directions);
        grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "ForecastProbability", HeaderText = "Probability (%)", FillWeight = 25 });
        root.Controls.Add(grid, 0, 1);
        var footer = ScenarioUi.Flow(); root.Controls.Add(footer, 0, 2);
        var missing = ScenarioUi.Button("Remove missing forecasts"); footer.Controls.Add(missing);
        footer.Controls.Add(new Label { AutoSize = true, Text = "Set the question per forecast. Each forecast uses the same settings across scenarios.", Margin = new Padding(3, 8, 3, 3) });
        missing.Click += (_, _) => {
            try { CaptureQuestions(); questions.RemoveAll(q => !forecasts.Any(f => ScenarioSimulationService.Identity(f).Equals(q.ForecastId, StringComparison.OrdinalIgnoreCase))); Render(); Changed?.Invoke(); }
            catch (ArgumentException ex) { MessageBox.Show(this, ex.Message, "Analysis Question", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        };
        mode.SelectedIndexChanged += (_, _) => {
            if (loading) return;
            try { CaptureQuestions(); currentMode = (ScenarioAnalysisMode)mode.SelectedIndex; Render(); Changed?.Invoke(); }
            catch (ArgumentException ex) { loading = true; mode.SelectedIndex = (int)currentMode; loading = false; MessageBox.Show(this, ex.Message, "Analysis Question", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        };

        grid.CellValueChanged += (_, _) => { if (!loading) Changed?.Invoke(); };
        grid.CurrentCellDirtyStateChanged += (_, _) => { if (grid.IsCurrentCellDirty) grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
        Render();
    }
    private void Render()
    {
        loading = true;
        try
        {
            grid.Rows.Clear(); grid.Columns[3].Visible = currentMode == ScenarioAnalysisMode.ProbabilityToValue;
            grid.Columns[1].Visible = currentMode == ScenarioAnalysisMode.ValueToProbability;
            foreach (var forecast in forecasts)
            {
                string id = ScenarioSimulationService.Identity(forecast);
                var q = questions.FirstOrDefault(x => x.ForecastId.Equals(id, StringComparison.OrdinalIgnoreCase)) ?? new(id);
                AddRow(q, ScenarioSimulationService.Display(forecast), false);
            }
            foreach (var q in questions.Where(q => !forecasts.Any(f => ScenarioSimulationService.Identity(f).Equals(q.ForecastId, StringComparison.OrdinalIgnoreCase))))
                AddRow(q, "Missing forecast: " + q.ForecastId, true);
        }
        finally { loading = false; }
    }
    private void AddRow(ScenarioForecastQuestion q, string display, bool missing)
    {
        var direction = currentMode == ScenarioAnalysisMode.ProbabilityToValue ? q.ValueDirection : q.ProbabilityDirection;
        int row = grid.Rows.Add(display, q.Target?.ToString("R", CultureInfo.CurrentCulture) ?? "",
            direction == TargetDirection.AtOrBelow ? "At or Below (≤)" : direction == TargetDirection.AtOrAbove ? "At or Above (≥)" : "Select direction", (q.Probability ?? legacyProbability).ToString("R", CultureInfo.CurrentCulture));
        grid.Rows[row].Tag = q.ForecastId;
        if (missing) { grid.Rows[row].ReadOnly = true; grid.Rows[row].ErrorText = "Remove this missing forecast before running."; }
    }
    private void CaptureQuestions()
    {
        if (!grid.EndEdit()) throw new ArgumentException("Finish editing the forecast question first.");
        var next = new List<ScenarioForecastQuestion>();
        foreach (DataGridViewRow row in grid.Rows)
        {
            string id = (string)row.Tag!;
            var q = questions.FirstOrDefault(q => q.ForecastId.Equals(id, StringComparison.OrdinalIgnoreCase)) ?? new(id);
            TargetDirection? direction = Convert.ToString(row.Cells[2].Value) switch
            { "At or Below (≤)" => TargetDirection.AtOrBelow, "At or Above (≥)" => TargetDirection.AtOrAbove, _ => null };
            if (currentMode == ScenarioAnalysisMode.ProbabilityToValue)
            {
                if (!double.TryParse(Convert.ToString(row.Cells[3].Value), out double p) || !double.IsFinite(p) || p < 0 || p > 100)
                    throw new ArgumentException($"{row.Cells[0].Value}: enter a probability from 0 to 100 percent.");
                q = q with { ValueDirection = direction, Probability = p };
            }
            else
            {
                string text = Convert.ToString(row.Cells[1].Value) ?? "";
                double? target = null;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    if (!double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out double value) || !double.IsFinite(value))
                        throw new ArgumentException($"{row.Cells[0].Value}: enter a finite target value.");
                    target = value;
                }
                q = q with { Target = target, ProbabilityDirection = direction };
            }
            next.Add(q);
        }
        questions = next;
    }
    public ScenarioAnalysisDefinition Snapshot(ScenarioAnalysisDefinition analysis)
    {
        CaptureQuestions();
        return analysis with { Mode = currentMode, ForecastQuestions = questions.ToArray() };
    }
    public void Reload(ScenarioAnalysisDefinition analysis)
    { questions = analysis.Questions.ToList(); Render(); }
}
