using System.Globalization;
using MonteCarlo.Core;
namespace MonteCarlo.Excel;

public sealed class OptimizerForm : Form
{
    private readonly IReadOnlyList<ForecastDefinition> forecasts;
    private OptimizationProblem problem;
    private readonly Func<OptimizationProblem, OptimizationProblem> save;
    private readonly Func<OptimizationProblem, CancellationToken, IProgress<OptimizationProgress>, Task<OptimizationResult>> run;
    private readonly Action<OptimizationResult> apply, export;
    private readonly ComboBox forecast = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260, DisplayMember = "Name" };
    private readonly ComboBox objective = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 250 };
    private readonly ComboBox direction = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 65 };
    private readonly TextBox target = new() { Width = 100 }, evaluations = new() { Width = 80 }, trials = new() { Width = 80 }, seed = new() { Width = 100 };
    private readonly DataGridView variables = ScenarioUi.Grid("DecisionVariables"), constraints = ScenarioUi.Grid("OptimizerConstraints");
    private readonly TextBox status = new() { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Height = 90, Text = "Configure controllable inputs. Evaluations update the model live; original values are restored afterward." };
    private readonly Button start = ScenarioUi.Button("Run Optimizer"), close = ScenarioUi.Button("Close"), cancel = ScenarioUi.Button("Cancel Optimization");
    private readonly TabControl tabs = new() { Dock = DockStyle.Fill };
    private CancellationTokenSource? cancellation;
    private readonly Func<DecisionVariableSelection?>? pickVariable;
    private readonly Func<DecisionVariable, DecisionVariableSelection>? readVariable;
    private readonly Action<IReadOnlyList<DecisionVariable>>? validateVariables;
    private readonly Func<DecisionVariableSelection?>? pickConstraint;
    private readonly Func<OptimizationConstraint, DecisionVariableSelection>? readConstraint;
    private readonly Action<OptimizationConstraint>? validateConstraint;
    private readonly bool deterministic;
    public OptimizerForm(IReadOnlyList<ForecastDefinition> forecasts, OptimizationProblem? initial,
        Func<OptimizationProblem, OptimizationProblem> save,
        Func<OptimizationProblem, CancellationToken, IProgress<OptimizationProgress>, Task<OptimizationResult>> run,
        Action<OptimizationResult> apply, Action<OptimizationResult> export,
        Func<DecisionVariableSelection?>? pickVariable = null, Func<DecisionVariable, DecisionVariableSelection>? readVariable = null,
        Action<IReadOnlyList<DecisionVariable>>? validateVariables = null,
        Func<DecisionVariableSelection?>? pickConstraint = null, Func<OptimizationConstraint, DecisionVariableSelection>? readConstraint = null,
        Action<OptimizationConstraint>? validateConstraint = null, bool deterministic = false)
    {
        this.forecasts = forecasts; this.save = save; this.run = run; this.apply = apply; this.export = export;
        this.deterministic = deterministic;
        this.pickVariable = pickVariable; this.readVariable = readVariable; this.validateVariables = validateVariables;
        this.pickConstraint = pickConstraint; this.readConstraint = readConstraint; this.validateConstraint = validateConstraint;
        problem = initial ?? new(new(forecasts.FirstOrDefault() is { } f ? ScenarioSimulationService.Identity(f) : ""), [], [], new());
        Text = "Optimizer"; AutoScaleMode = AutoScaleMode.Dpi; ClientSize = new(1000, 650); MinimumSize = new(800, 500); StartPosition = FormStartPosition.CenterParent;
        var root = ScenarioUi.Table(1, 3); root.Padding = new(12); Controls.Add(root);
        root.RowStyles.Add(new(SizeType.AutoSize)); root.RowStyles.Add(new(SizeType.Percent, 100)); root.RowStyles.Add(new(SizeType.AutoSize));
        root.Controls.Add(status, 0, 0); root.Controls.Add(tabs, 0, 1);
        var objectivePage = new TabPage("Objective / Settings"); var variablePage = new TabPage("Decision Variables"); var constraintPage = new TabPage("Constraints");
        tabs.TabPages.AddRange([objectivePage, variablePage, constraintPage]);
        var settings = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, Padding = new(10) };
        objectivePage.Controls.Add(settings);
        forecast.Items.AddRange(forecasts.Cast<object>().ToArray());
        forecast.FormattingEnabled = true;
        forecast.Format += (_, e) => { if (e.ListItem is ForecastDefinition f) e.Value = ScenarioSimulationService.Display(f); };
        forecast.SelectedItem = forecasts.FirstOrDefault(f => ScenarioSimulationService.Identity(f) == problem.Objective.ForecastId);
        objective.Items.AddRange(["Maximize Mean", "Minimize Mean", "Maximize Median (P50)", "Minimize Median (P50)", "Maximize Probability of Target", "Minimize Probability of Target"]);
        if (deterministic) { objective.Items.RemoveAt(5); objective.Items.RemoveAt(4); }
        objective.SelectedIndex = deterministic && problem.Objective.Metric == OptimizationMetric.Probability ? -1 : (int)problem.Objective.Metric * 2 + (problem.Objective.Maximize ? 0 : 1);
        direction.Items.AddRange([">=", "<="]); direction.SelectedIndex = problem.Objective.Direction == TargetDirection.AtOrAbove ? 0 : 1;
        target.Text = F(problem.Objective.Target); evaluations.Text = problem.Settings.MaximumEvaluations.ToString(); trials.Text = problem.Settings.Trials.ToString(); seed.Text = problem.Settings.Seed.ToString();
        settings.Controls.Add(Row("Forecast", forecast)); settings.Controls.Add(Row("Optimize", objective));
        settings.Controls.Add(Row("Target Value / Direction", target, direction));
        objective.SelectedIndexChanged += (_, _) => { target.Enabled = direction.Enabled = objective.SelectedIndex >= 4; };
        target.Enabled = direction.Enabled = objective.SelectedIndex >= 4;
        settings.Controls.Add(Row("Search Method", new Label { Text = "Automatic", AutoSize = true }));
        settings.Controls.Add(Row("Evaluation Mode", new Label { Name = "EvaluationMode", Text = deterministic ? "Deterministic" : "Monte Carlo", AutoSize = true }));
        trials.Enabled = !deterministic; trials.Name = "TrialsPerEvaluation";
        settings.Controls.Add(Row("Maximum Candidate Evaluations", evaluations));
        settings.Controls.Add(new Label { AutoSize = true, MaximumSize = new(650, 0), Text = "Maximum number of decision-variable combinations the optimizer may evaluate." });
        if (deterministic) settings.Controls.Add(Row("Trials per Evaluation", new Label { Text = "N/A – Deterministic model", AutoSize = true }));
        else settings.Controls.Add(Row("Monte Carlo Trials per Evaluation", trials));
        if (!deterministic) settings.Controls.Add(new Label { AutoSize = true, MaximumSize = new(650, 0), Text = "Number of Monte Carlo simulation trials used to estimate the forecast distribution for each candidate solution." });
        var estimate = new Label { Name = "ComputationalEstimate", AutoSize = true, MaximumSize = new(650, 0) };
        void Estimate()
        {
            estimate.Text = int.TryParse(evaluations.Text, out int e) && int.TryParse(trials.Text, out int t) && e > 0 && t > 0
                ? deterministic ? $"Candidate Evaluations: {e:N0}; one recalculation per candidate. Maximum Monte Carlo Trials: 0."
                : $"Candidate Evaluations: {e:N0}   Trials per Candidate: {t:N0}\nMaximum Monte Carlo Trials: {(long)e * t:N0} (upper bound; rejected candidates skip simulation)."
                : "Enter valid settings to estimate work.";
        }
        evaluations.TextChanged += (_, _) => Estimate(); trials.TextChanged += (_, _) => Estimate(); Estimate(); settings.Controls.Add(estimate);
        settings.Controls.Add(Row(deterministic ? "Search Random Seed" : "Random Seed", seed));
        settings.Controls.Add(new Label { AutoSize = true, MaximumSize = new(650, 0), Text = deterministic
            ? "One model recalculation per candidate. Mean and P50 equal the calculated forecast value. Probability optimization is unavailable without uncertainty. Step anchors the grid at Minimum; the search seed controls exploration."
            : "Step defines a numeric grid anchored at Minimum. Maximum is included only when aligned with Step. The common seed is reused for every candidate. A larger budget improves coverage; results remain Monte Carlo estimates." });
        SetupGrid(variablePage, variables, ["Variable", "Cell", "Minimum", "Maximum", "Step"], EditVariable, () =>
        { if (Index(variables) is int i) { problem = problem with { Variables = problem.Variables.Where((_, n) => n != i).ToArray() }; RefreshGrids(); } });
        SetupGrid(constraintPage, constraints, ["Type", "Cell / Forecast / Variable", "Condition", "Value", "Requirement"], EditConstraint, () =>
        { if (Index(constraints) is int i) { problem = problem with { Constraints = problem.Constraints.Where((_, n) => n != i).ToArray() }; RefreshGrids(); } });
        var footer = ScenarioUi.Flow(); root.Controls.Add(footer, 0, 2); var saveButton = ScenarioUi.Button("Save Configuration");
        footer.Controls.AddRange([saveButton, start, cancel, close]); cancel.Enabled = false;
        saveButton.Click += (_, _) => { if (cancellation == null) Guard(() => { problem = save(Read()); RefreshGrids(); status.Text = "Configuration saved in workbook. Use Excel Save to retain it after reopening."; }); };
        close.Click += (_, _) => Close(); cancel.Click += (_, _) => cancellation?.Cancel(); start.Click += async (_, _) => await Run();
        FormClosing += (_, e) => { if (cancellation != null) { e.Cancel = true; cancellation.Cancel(); status.Text = "Cancelling; waiting for safe cleanup…"; } };
        RefreshGrids();
    }
    private static string F(double value) => value.ToString("G17", CultureInfo.CurrentCulture);
    private static FlowLayoutPanel Row(string label, params Control[] controls)
    {
        var row = ScenarioUi.Flow(); row.Controls.Add(new Label { Text = label, Width = 240, AutoSize = false, Height = 28 }); row.Controls.AddRange(controls); return row;
    }
    private static int? Index(DataGridView grid) => grid.CurrentRow?.Index;
    private void Guard(Action action) { try { action(); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Optimizer", MessageBoxButtons.OK, MessageBoxIcon.Warning); } }
    private void SetupGrid(TabPage page, DataGridView grid, string[] headers, Action<bool> edit, Action delete)
    {
        grid.ReadOnly = true; grid.MultiSelect = false; grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        foreach (var header in headers) grid.Columns.Add(header, header);
        var root = ScenarioUi.Table(1, 2); root.RowStyles.Add(new(SizeType.Percent, 100)); root.RowStyles.Add(new(SizeType.AutoSize)); page.Controls.Add(root); root.Controls.Add(grid, 0, 0);
        var buttons = ScenarioUi.Flow(); root.Controls.Add(buttons, 0, 1);
        foreach (var label in new[] { "Add", "Edit", "Delete" }) { var button = ScenarioUi.Button(label); buttons.Controls.Add(button); button.Click += (_, _) => Guard(() => { if (label == "Delete") delete(); else edit(label == "Edit"); }); }
    }
    private OptimizationProblem Read()
    {
        if (forecast.SelectedItem is not ForecastDefinition f) throw new ArgumentException("Objective: select a configured forecast. Define a forecast first if the list is empty.");
        if (objective.SelectedIndex < 0) throw new ArgumentException("Deterministic optimization: choose Mean or Median. Probability objectives require assumptions.");
        if (!int.TryParse(evaluations.Text, out int e) || !int.TryParse(trials.Text, out int t) || !int.TryParse(seed.Text, out int s)) throw new ArgumentException("Settings: enter integer evaluations, trials and seed.");
        var result = problem with { Objective = new(ScenarioSimulationService.Identity(f), (OptimizationMetric)(objective.SelectedIndex / 2), objective.SelectedIndex % 2 == 0,
            objective.SelectedIndex >= 4 ? Number(target.Text, "Objective target") : 0, direction.SelectedIndex == 0 ? TargetDirection.AtOrAbove : TargetDirection.AtOrBelow), Settings = new(e, t, s) };
        result.Validate(deterministic); OptimizationService.ValidateMode(result, deterministic); return result;
    }
    private static double Number(string text, string label) => double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out double value) && double.IsFinite(value)
        ? value : throw new ArgumentException($"{label}: enter a finite numeric value.");
    private void RefreshGrids()
    {
        variables.Rows.Clear(); foreach (var v in problem.Variables) variables.Rows.Add(v.Name, v.Sheet + "!" + v.Cell, v.Minimum, v.Maximum, v.Step);
        constraints.Rows.Clear(); foreach (var c in problem.Constraints) constraints.Rows.Add(c.Kind, c.Kind == OptimizationConstraintKind.Linear ? OptimizationFeasibleSpace.Expression(c, problem.Variables) : c.Kind == OptimizationConstraintKind.Cell ? c.CellDescription : c.Kind == OptimizationConstraintKind.DecisionVariable ? problem.Variables.FirstOrDefault(v => v.Id == c.SubjectId)?.Name ?? "Missing variable" : forecasts.FirstOrDefault(f => ScenarioSimulationService.Identity(f) == c.SubjectId)?.Name ?? "Missing forecast",
            c.Kind == OptimizationConstraintKind.Probability ? $"P(value {OptimizerExporter.Direction(c.TargetDirection)} {c.Target:G8})" : OptimizerExporter.Operator(c),
            c.Value, c.Kind == OptimizationConstraintKind.Probability ? $"{OptimizerExporter.Operator(c)} {c.Value:G8}%" : "");
    }
    private void EditVariable(bool editing)
    {
        int index = editing ? Index(variables) ?? -1 : -1; if (editing && index < 0) return;
        var old = index >= 0 ? problem.Variables[index] : null;
        using var dialog = new DecisionVariableForm(old, () =>
        {
            try { Hide(); return pickVariable?.Invoke(); }
            finally { Show(); Activate(); }
        }, v => readVariable?.Invoke(v) ?? throw new ArgumentException("Cell selection is unavailable."), v =>
        {
            var list = problem.Variables.ToList(); if (index < 0) list.Add(v); else list[index] = v;
            validateVariables?.Invoke(list);
            problem = problem with { Variables = list };
        });
        if (dialog.ShowDialog(this) == DialogResult.OK) RefreshGrids();
    }
    private void EditConstraint(bool editing)
    {
        int index = editing ? Index(constraints) ?? -1 : -1; if (editing && index < 0) return;
        var old = index >= 0 ? problem.Constraints[index] : null;
        using var dialog = new OptimizerConstraintForm(old, problem.Variables, forecasts, () =>
        {
            try { Hide(); return pickConstraint?.Invoke(); }
            finally { Show(); Activate(); }
        }, c => readConstraint?.Invoke(c) ?? throw new ArgumentException("Constraint cell selection is unavailable."), validateConstraint, deterministic);
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var list = problem.Constraints.ToList(); if (index < 0) list.Add(dialog.Constraint!); else list[index] = dialog.Constraint!;
        problem = problem with { Constraints = list }; RefreshGrids();
    }
    private async Task Run()
    {
        try
        {
            problem = save(Read()); RefreshGrids(); cancellation = new(); tabs.Enabled = false; start.Enabled = close.Enabled = false; cancel.Enabled = true;
            var result = await run(problem, cancellation.Token, new Progress<OptimizationProgress>(p =>
            {
                status.Text = OptimizationProgressText.Format(p, problem, deterministic);
            }));
            status.Text = "Optimization completed. Original model inputs restored.";
            using var results = new OptimizerResultsForm(result, apply, export); results.ShowDialog(this);
        }
        catch (OperationCanceledException) { status.Text = "Optimization cancelled. Original model inputs restored."; }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Optimizer", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        finally { cancellation?.Dispose(); cancellation = null; tabs.Enabled = true; start.Enabled = close.Enabled = true; cancel.Enabled = false; }
    }
}

internal sealed class OptimizerEditor : Form
{
    private readonly Control[] fields;
    public string[] Values => fields.Select(c => c.Text.Trim()).ToArray();
    public OptimizerEditor(string title, string[] labels, string[] initial, Dictionary<int, string[]>? choices = null)
    {
        Text = title; AutoScaleMode = AutoScaleMode.Dpi; ClientSize = new(640, 360); MinimumSize = new(600, 400); StartPosition = FormStartPosition.CenterParent;
        var root = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, Padding = new(12) }; Controls.Add(root);
        fields = new Control[labels.Length];
        for (int i = 0; i < labels.Length; i++)
        {
            if (choices != null && choices.TryGetValue(i, out var options)) { var combo = new ComboBox { Width = 320, DropDownStyle = ComboBoxStyle.DropDownList }; combo.Items.AddRange(options); combo.SelectedItem = initial[i]; fields[i] = combo; }
            else fields[i] = new TextBox { Text = initial[i], Width = 320 };
            var row = ScenarioUi.Flow(); row.Controls.Add(new Label { Text = labels[i], Width = 245, Height = 27 }); row.Controls.Add(fields[i]); root.Controls.Add(row);
        }
        var footer = ScenarioUi.Flow(); var ok = ScenarioUi.Button("OK"); var cancel = ScenarioUi.Button("Cancel"); ok.DialogResult = DialogResult.OK; cancel.DialogResult = DialogResult.Cancel; footer.Controls.AddRange([ok, cancel]); root.Controls.Add(footer); AcceptButton = ok; CancelButton = cancel;
    }
}

public sealed class OptimizerResultsForm : Form
{
    public OptimizerResultsForm(OptimizationResult result, Action<OptimizationResult> apply, Action<OptimizationResult> export)
    {
        Text = "Optimizer Results"; AutoScaleMode = AutoScaleMode.Dpi; ClientSize = new(1080, 680); MinimumSize = new(800, 500); StartPosition = FormStartPosition.CenterParent;
        var root = ScenarioUi.Table(1, 4); root.Padding = new(12); Controls.Add(root); root.RowStyles.Add(new(SizeType.AutoSize)); root.RowStyles.Add(new(SizeType.AutoSize)); root.RowStyles.Add(new(SizeType.Percent, 100)); root.RowStyles.Add(new(SizeType.AutoSize));
        root.Controls.Add(new Label { AutoSize = true, Dock = DockStyle.Fill, Text = result.Best.Feasible ? "OPTIMAL SOLUTION — best feasible candidate found" : "NO FEASIBLE SOLUTION — adjust constraints or increase evaluations" }, 0, 0);
        var p = result.Problem; var forecasts = result.ConfiguredForecasts.Count > 0 ? result.ConfiguredForecasts : result.Best.Run.ForecastResults.Select(f => f.Forecast).ToArray();
        bool deterministic = result.EvaluationMode == OptimizationEvaluationMode.Deterministic;
        root.Controls.Add(new TextBox { Dock = DockStyle.Fill, ReadOnly = true, Multiline = true, ScrollBars = ScrollBars.Vertical, Height = 135,
            Text = $"Objective: {OptimizerExporter.Objective(p, forecasts)}\r\nBest Feasible Objective: {(result.BestFeasible?.Objective.ToString("N6") ?? "Unavailable")}{(p.Objective.Metric == OptimizationMetric.Probability ? "%" : "")}\r\n" +
                $"Evaluation Mode: {(deterministic ? "Deterministic" : "Monte Carlo")}\r\nEvaluations: {result.Evaluations} / {p.Settings.MaximumEvaluations}    Trials per Evaluation: {(deterministic ? "N/A" : p.Settings.Trials.ToString("N0"))}    Search Seed: {p.Settings.Seed}\r\nDuration: {result.Duration.TotalSeconds:N2} seconds    Search Method: Automatic\r\n" +
                $"Rejected before simulation: {result.RejectedBeforeSimulation:N0}; feasible evaluated: {result.FeasibleEvaluations:N0}; Monte Carlo evaluations: {result.MonteCarloEvaluations:N0}\r\n" +
                (deterministic ? "Original inputs retained until Apply. Calculated forecast values; no global optimality guarantee." : "Original inputs retained until Apply. Results are Monte Carlo estimates; no global optimality guarantee.") }, 0, 1);
        var tabs = new TabControl { Dock = DockStyle.Fill }; root.Controls.Add(tabs, 0, 2);
        DataGridView Grid(string title, string[] headers)
        {
            var page = new TabPage(title); tabs.TabPages.Add(page); var grid = ScenarioUi.Grid("Optimizer" + title.Replace(" ", "")); grid.ReadOnly = true;
            grid.DefaultCellStyle.WrapMode = DataGridViewTriState.True; grid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            foreach (var h in headers) grid.Columns.Add(h, h); grid.Columns[0].FillWeight = 220; page.Controls.Add(grid); return grid;
        }
        var variables = Grid("Decision Variables", ["Variable", "Cell", "Original", result.BestFeasible != null ? "Optimized" : "Infeasible diagnostic candidate", "Change"]);
        for (int i = 0; i < p.Variables.Count; i++) { var v = p.Variables[i]; double change = result.Best.Values[i] - result.OriginalValues[i]; variables.Rows.Add(v.Name, v.Sheet + "!" + v.Cell, result.OriginalValues[i], result.Best.Values[i], change.ToString("+0.######;-0.######;0")); }
        var constraints = Grid("Constraints", ["Constraint", "Result", "Status"]); constraints.Columns[0].FillWeight = 400;
        foreach (var c in result.Best.Constraints) constraints.Rows.Add(OptimizerExporter.Constraint(c.Constraint, p, forecasts), c.Actual.ToString("N6") + (c.Constraint.Kind == OptimizationConstraintKind.Probability ? "%" : ""), c.Met ? "Met" : "Not met");
        var forecastGrid = Grid("Forecast Results", ["Forecast", "Mean", "P50", "P80", "Standard Deviation", "Target Probability", "Target", "Direction"]);
        foreach (var f in result.Best.Run.ForecastResults)
        {
            bool selected = p.Objective.Metric == OptimizationMetric.Probability && ScenarioSimulationService.Identity(f.Forecast) == p.Objective.ForecastId;
            double? target = selected ? p.Objective.Target : f.Forecast.TargetSettings?.Target; var d = selected ? p.Objective.Direction : f.Forecast.TargetSettings?.Direction ?? MonteCarlo.Core.TargetDirection.AtOrAbove;
            forecastGrid.Rows.Add(ScenarioSimulationService.Display(f.Forecast), f.Mean, f.P50, f.P80, f.StandardDeviation,
                !deterministic && target.HasValue ? OptimizationService.Statistic(f, OptimizationMetric.Probability, target.Value, d).ToString("N2") + "%" : "—", target?.ToString("G8") ?? "—", target.HasValue ? OptimizerExporter.Direction(d) : "—");
        }
        var footer = ScenarioUi.Flow(); root.Controls.Add(footer, 0, 3); var applyButton = ScenarioUi.Button("Apply Optimal Values"); var exportButton = ScenarioUi.Button("Export to Excel"); var close = ScenarioUi.Button("Close Without Applying"); footer.Controls.AddRange([applyButton, exportButton, close]); applyButton.Enabled = result.Best.Feasible;
        applyButton.Click += (_, _) => { try { apply(result); applyButton.Enabled = false; close.Text = "Close"; MessageBox.Show(this, "Optimal values applied to the source workbook. Use Excel Save to retain them.", "Optimizer"); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Apply Optimal Values", MessageBoxButtons.OK, MessageBoxIcon.Warning); } };
        exportButton.Click += (_, _) => { try { export(result); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Optimizer Export", MessageBoxButtons.OK, MessageBoxIcon.Warning); } }; close.Click += (_, _) => Close();
    }
}
