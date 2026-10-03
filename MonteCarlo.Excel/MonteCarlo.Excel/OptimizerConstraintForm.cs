using System.Globalization;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public sealed class OptimizerConstraintForm : Form
{
    private readonly ComboBox kind = new() { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill, Name = "ConstraintType" };
    private readonly ComboBox subject = new() { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill, Name = "ConstraintSubject" };
    private readonly ComboBox comparison = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 70, Name = "ConstraintOperator" };
    private readonly ComboBox targetDirection = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 70 };
    private readonly TextBox reference = new() { Dock = DockStyle.Fill, Name = "ConstraintReference" }, name = new() { Dock = DockStyle.Fill, Name = "ConstraintName" };
    private readonly TextBox rhs = new() { Width = 180, Name = "ConstraintValue" }, target = new() { Width = 180 };
    private readonly Label current = new() { AutoSize = true, Text = "Select or type a numeric cell", Name = "ConstraintCurrentValue" };
    private readonly TextBox expression = new() { Dock = DockStyle.Fill, Name = "LinearExpression" };
    private DecisionVariableSelection? original;
    public OptimizationConstraint? Constraint { get; private set; }
    public OptimizerConstraintForm(OptimizationConstraint? existing, IReadOnlyList<DecisionVariable> variables, IReadOnlyList<ForecastDefinition> forecasts,
        Func<DecisionVariableSelection?> pick, Func<OptimizationConstraint, DecisionVariableSelection> read, Action<OptimizationConstraint>? validate = null, bool deterministic = false)
    {
        var kinds = new List<OptimizationConstraintKind> { OptimizationConstraintKind.Cell, OptimizationConstraintKind.DecisionVariable, OptimizationConstraintKind.Mean, OptimizationConstraintKind.Median };
        if (!deterministic) kinds.Add(OptimizationConstraintKind.Probability);
        kinds.Add(OptimizationConstraintKind.Linear);
        Text = "Optimizer Constraint"; AutoScaleMode = AutoScaleMode.Dpi; ClientSize = new(610, 410); MinimumSize = new(560, 390); StartPosition = FormStartPosition.CenterParent;
        var layout = ScenarioUi.Table(1, 2); layout.RowStyles.Add(new(SizeType.Percent, 100)); layout.RowStyles.Add(new(SizeType.AutoSize)); Controls.Add(layout);
        var body = new Panel { Dock = DockStyle.Fill, AutoScroll = true }; layout.Controls.Add(body, 0, 0);
        var root = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new(12), ColumnCount = 2, RowCount = 9 };
        root.ColumnStyles.Add(new(SizeType.AutoSize)); root.ColumnStyles.Add(new(SizeType.Percent, 100)); body.Controls.Add(root);
        for (int i = 0; i < 9; i++) root.RowStyles.Add(new(SizeType.AutoSize));
        var labels = new Dictionary<int, Label>();
        void Row(int row, string label, Control control)
        { var caption = new Label { Text = label, AutoSize = true, Margin = new(3, 7, 12, 7) }; labels[row] = caption; root.Controls.Add(caption, 0, row); root.Controls.Add(control, 1, row); }
        Row(0, "Constraint", kind); Row(1, "Display Name", name);
        var referenceRow = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2, RowCount = 1 };
        referenceRow.ColumnStyles.Add(new(SizeType.Percent, 100)); referenceRow.ColumnStyles.Add(new(SizeType.AutoSize));
        var select = ScenarioUi.Button("Select"); select.Name = "SelectConstraintCell"; referenceRow.Controls.Add(reference, 0, 0); referenceRow.Controls.Add(select, 1, 0);
        Row(2, "Excel Cell", referenceRow); Row(3, "Forecast / Variable", subject);
        var condition = ScenarioUi.Flow(); condition.Controls.AddRange([comparison, rhs]); Row(4, "Condition / RHS", condition);
        Row(5, "Current Value", current);
        var probability = ScenarioUi.Flow(); probability.Controls.AddRange([targetDirection, target]); Row(6, "Probability Target", probability);
        var help = new Label { AutoSize = true, MaximumSize = new(420, 0), Text = "Any numeric value or formula cell. Enter Sheet1!$B$6 or select in Excel. Cells are read after recalculation for every candidate; no Forecast is required. Equality uses a small numeric tolerance." };
        Row(7, "Linear Expression", expression); root.Controls.Add(help, 1, 8);
        var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new(9, 3, 9, 9) };
        var ok = ScenarioUi.Button("OK"); var cancel = ScenarioUi.Button("Cancel"); cancel.DialogResult = DialogResult.Cancel;
        footer.Controls.AddRange([ok, cancel]); layout.Controls.Add(footer, 0, 1); AcceptButton = ok; CancelButton = cancel;
        kind.Items.AddRange(kinds.Select(k => (object)(k == OptimizationConstraintKind.Linear ? "Linear decision expression" : k == OptimizationConstraintKind.Cell ? "Excel cell (value / formula)" : k == OptimizationConstraintKind.DecisionVariable ? "Decision variable" : "Forecast " + k.ToString().ToLowerInvariant())).ToArray());
        comparison.Items.AddRange(["<=", ">=", "="]); comparison.SelectedItem = existing == null ? "<=" : OptimizerExporter.Operator(existing);
        targetDirection.Items.AddRange([">=", "<="]); targetDirection.SelectedIndex = existing?.TargetDirection == TargetDirection.AtOrBelow ? 1 : 0;
        rhs.Text = (existing?.Value ?? 0).ToString("G17", CultureInfo.CurrentCulture); target.Text = (existing?.Target ?? 0).ToString("G17", CultureInfo.CurrentCulture); name.Text = existing?.Name ?? "";
        var subjects = new List<(string Id, string Label)>();
        OptimizationConstraintKind Kind() => kinds[kind.SelectedIndex];
        kind.SelectedIndexChanged += (_, _) =>
        {
            bool cell = Kind() == OptimizationConstraintKind.Cell, prob = Kind() == OptimizationConstraintKind.Probability, linear = Kind() == OptimizationConstraintKind.Linear;
            name.Visible = labels[1].Visible = cell || linear;
            referenceRow.Visible = labels[2].Visible = current.Visible = labels[5].Visible = cell;
            subject.Visible = labels[3].Visible = !cell && !linear; probability.Visible = labels[6].Visible = prob;
            expression.Visible = labels[7].Visible = linear;
            labels[4].Text = prob ? "Requirement (%)" : "Condition / RHS";
            subjects = Kind() == OptimizationConstraintKind.DecisionVariable
                ? variables.Select(v => (v.Id, $"{v.Name} ({v.Sheet}!{v.Cell})")).ToList()
                : forecasts.Select(f => (ScenarioSimulationService.Identity(f), ScenarioSimulationService.Display(f))).ToList();
            subject.Items.Clear(); subject.Items.AddRange(subjects.Select(x => (object)x.Label).ToArray());
            subject.SelectedIndex = subjects.FindIndex(x => x.Id == existing?.SubjectId);
            if (subject.SelectedIndex < 0 && subjects.Count > 0) subject.SelectedIndex = 0;
            help.Text = linear ? "Use decision-variable names: X + Y + Z or 2*X + Y. Coefficients use a decimal point. Bounds and Step remain enforced; equality uses relative tolerance 1e-9. Infeasible candidates skip simulation."
                : cell ? "Any numeric value or formula cell. Enter Sheet1!$B$6 or select in Excel. Cells are read after recalculation for every candidate; no Forecast is required. Equality uses a small numeric tolerance."
                : "Existing decision-variable and Monte Carlo forecast-statistic constraints are supported. Equality uses a small numeric tolerance.";
        };
        kind.SelectedIndex = deterministic && existing?.Kind == OptimizationConstraintKind.Probability ? 2 : kinds.IndexOf(existing?.Kind ?? OptimizationConstraintKind.Cell);
        if (existing?.Kind == OptimizationConstraintKind.Linear) expression.Text = OptimizationFeasibleSpace.Expression(existing, variables);
        void Set(DecisionVariableSelection selected)
        { reference.Text = ConstraintCellSelection.Reference(selected.Sheet, selected.Cell); current.Text = selected.CurrentValue.ToString("G17", CultureInfo.CurrentCulture); name.Text = DecisionVariableSelection.SuggestName(name.Text, selected.SuggestedName); }
        if (existing?.Kind == OptimizationConstraintKind.Cell)
        {
            reference.Text = existing.Sheet + "!" + existing.Cell;
            try { original = read(existing); Set(original); } catch (Exception ex) { current.Text = "Reference unavailable — replace it"; current.Tag = ex.Message; }
        }
        reference.TextChanged += (_, _) => current.Text = "Validated when you click OK";
        select.Click += (_, _) =>
        {
            try { Hide(); var selected = pick(); if (selected != null) Set(selected); }
            catch (Exception ex) { MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { Show(); Activate(); }
        };
        ok.Click += (_, _) =>
        {
            try
            {
                double Number(TextBox box, string label) => double.TryParse(box.Text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out var v) && double.IsFinite(v) ? v : throw new ArgumentException($"{label}: enter a finite numeric value.");
                var op = comparison.SelectedIndex switch { 0 => OptimizationConstraintOperator.LessThanOrEqual, 1 => OptimizationConstraintOperator.GreaterThanOrEqual, _ => OptimizationConstraintOperator.Equal };
                var draft = new OptimizationConstraint(Kind(), existing?.SubjectId ?? Guid.NewGuid().ToString("N"), op == OptimizationConstraintOperator.LessThanOrEqual ? TargetDirection.AtOrBelow : TargetDirection.AtOrAbove,
                    Number(rhs, "RHS"), Kind() == OptimizationConstraintKind.Probability ? Number(target, "Probability target") : 0,
                    targetDirection.SelectedIndex == 1 ? TargetDirection.AtOrBelow : TargetDirection.AtOrAbove, Operator: op);
                if (Kind() == OptimizationConstraintKind.Cell)
                {
                    var parsed = ConstraintCellSelection.Parse(reference.Text);
                    draft = draft with { Sheet = parsed.Sheet, Cell = parsed.Cell, Name = name.Text.Trim() };
                    var selected = read(draft);
                    if (original != null && original.Sheet.Equals(selected.Sheet, StringComparison.OrdinalIgnoreCase) && original.Cell == selected.Cell)
                        draft = draft with { CellLink = existing!.CellLink };
                }
                else if (Kind() == OptimizationConstraintKind.Linear)
                    draft = draft with { Terms = OptimizationFeasibleSpace.Parse(expression.Text, variables), Name = name.Text.Trim() };
                else
                {
                    if (subject.SelectedIndex < 0) throw new ArgumentException("Select a configured forecast or decision variable.");
                    draft = draft with { SubjectId = subjects[subject.SelectedIndex].Id };
                }
                if (draft.Kind == OptimizationConstraintKind.Probability && (draft.Value < 0 || draft.Value > 100)) throw new ArgumentException("Probability requirement must be between 0 and 100 percent.");
                validate?.Invoke(draft); Constraint = draft; DialogResult = DialogResult.OK; Close();
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        };
    }
}
