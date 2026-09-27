using System.Globalization;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public sealed class ScenarioAnalysisForm : Form
{
    private readonly IReadOnlyList<AssumptionDefinition> assumptions;
    private readonly Func<ScenarioAnalysisDefinition, ScenarioAnalysisDefinition> save;
    private readonly Func<ScenarioAnalysisDefinition, CancellationToken, IProgress<string>, Task<ScenarioComparisonResult>> run;
    private readonly IReadOnlyList<ForecastDefinition> forecasts;
    private ScenarioAnalysisDefinition analysis;
    private readonly ScenarioQuestionControl question;
    private readonly Label baselineLabel = new() { AutoSize = true, Dock = DockStyle.Fill, Name = "BaselineSelection" };
    private List<ScenarioDefinition> scenarios;
    private readonly CheckedListBox list = new() { Dock = DockStyle.Fill, CheckOnClick = false, DisplayMember = "Name", Name = "Scenarios" };
    private readonly TextBox name = new() { Width = 190, Name = "ScenarioName" };
    private readonly DataGridView grid = ScenarioUi.Grid("Adjustments");
    private readonly Label status = new() { AutoSize = true, Dock = DockStyle.Fill, Text = "Check scenarios to compare. Blank change means No change; enter 15 for +15%.", Name = "ScenarioStatus" };
    private readonly Button runButton = ScenarioUi.Button("Run Scenario Analysis"), close = ScenarioUi.Button("Close");
    private readonly Control editor;
    private CancellationTokenSource? cancellation;
    private bool loading, dirty;

    public ScenarioAnalysisForm(IReadOnlyList<AssumptionDefinition> assumptions, IReadOnlyList<ForecastDefinition> forecasts, ScenarioAnalysisDefinition initialAnalysis,
        Func<ScenarioAnalysisDefinition, ScenarioAnalysisDefinition> save,
        Func<ScenarioAnalysisDefinition, CancellationToken, IProgress<string>, Task<ScenarioComparisonResult>> run)
    {
        this.assumptions = assumptions; this.forecasts = forecasts; analysis = initialAnalysis;
        scenarios = analysis.Scenarios.ToList(); this.save = save; this.run = run;
        Text = "Scenario Analysis"; AutoScaleMode = AutoScaleMode.Dpi; ClientSize = new(980, 590);
        MinimumSize = new(760, 440); StartPosition = FormStartPosition.CenterParent;
        var root = ScenarioUi.Table(1, 4); root.Padding = new Padding(12); Controls.Add(root);
        root.RowStyles.Add(new(SizeType.AutoSize)); root.RowStyles.Add(new(SizeType.AutoSize)); root.RowStyles.Add(new(SizeType.Percent, 100)); root.RowStyles.Add(new(SizeType.AutoSize));
        root.Controls.Add(status, 0, 0);
        root.Controls.Add(baselineLabel, 0, 1);
        var tabs = new TabControl { Dock = DockStyle.Fill, Name = "ScenarioTabs" };
        var scenariosPage = new TabPage("Scenarios"); var questionPage = new TabPage("Analysis Question");
        tabs.TabPages.AddRange(new[] { scenariosPage, questionPage }); root.Controls.Add(tabs, 0, 2);
        question = new ScenarioQuestionControl(forecasts, analysis); question.Changed += () => dirty = true; questionPage.Controls.Add(question);
        var body = ScenarioUi.Table(2, 1); body.ColumnStyles.Add(new(SizeType.Percent, 27)); body.ColumnStyles.Add(new(SizeType.Percent, 73));
        scenariosPage.Controls.Add(body); editor = tabs;
        body.Controls.Add(list, 0, 0);
        var right = ScenarioUi.Table(1, 2); right.RowStyles.Add(new(SizeType.AutoSize)); right.RowStyles.Add(new(SizeType.Percent, 100));
        body.Controls.Add(right, 1, 0);
        var names = ScenarioUi.Flow(); right.Controls.Add(names, 0, 0);
        names.Controls.Add(new Label { Text = "Scenario name", AutoSize = true, Margin = new Padding(3, 8, 3, 3) }); names.Controls.Add(name);
        var add = ScenarioUi.Button("New"); var rename = ScenarioUi.Button("Rename"); var delete = ScenarioUi.Button("Delete");
        names.Controls.AddRange(new Control[] { add, rename, delete });
        var baseline = ScenarioUi.Button("Set as Baseline"); names.Controls.Add(baseline);
        baseline.Click += (_, _) => { if (Selected is { } selected) { analysis = analysis with { BaselineScenarioId = selected.Id }; dirty = true; RefreshList(selected.Id); } };
        list.FormattingEnabled = true;
        list.Format += (_, e) => { if (e.ListItem is ScenarioDefinition s) e.Value = s.Name + (s.Id == analysis.BaselineScenarioId ? " ★ Baseline" : ""); };
        list.ItemCheck += (_, e) => { if (list.Items[e.Index] is ScenarioDefinition s && s.Id == analysis.BaselineScenarioId) e.NewValue = CheckState.Checked; };
        right.Controls.Add(grid, 0, 1);
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Assumption", ReadOnly = true, FillWeight = 45 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Distribution", ReadOnly = true, FillWeight = 25 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Change (%)", FillWeight = 30 });
        grid.Columns[2].DefaultCellStyle.NullValue = "No change";
        grid.CellFormatting += (_, e) => {
            if (e.ColumnIndex == 2 && string.IsNullOrWhiteSpace(Convert.ToString(e.Value)))
            { e.Value = "No change"; e.FormattingApplied = true; }
        };
        var footer = ScenarioUi.Flow(); root.Controls.Add(footer, 0, 3);
        var remove = ScenarioUi.Button("No change"); var saveButton = ScenarioUi.Button("Save definitions");
        footer.Controls.AddRange(new Control[] { remove, saveButton, runButton, close });
        list.SelectedIndexChanged += (_, _) => ShowScenario();
        add.Click += (_, _) => Guard(() => {
            var next = new ScenarioDefinition(Guid.NewGuid().ToString("N"), name.Text.Trim(), Array.Empty<ScenarioAdjustment>());
            ScenarioValidation.Validate(scenarios.Append(next).ToArray()); scenarios.Add(next); dirty = true; RefreshList(next.Id);
        });
        rename.Click += (_, _) => Guard(() => {
            if (Selected is not { } selected) return;
            var next = scenarios.Select(s => s.Id == selected.Id ? s with { Name = name.Text.Trim() } : s).ToList();
            ScenarioValidation.Validate(next); scenarios = next; dirty = true; RefreshList(selected.Id);
        });
        delete.Click += (_, _) => {
            if (Selected is not { } selected || MessageBox.Show(this, $"Delete scenario '{selected.Name}'?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            scenarios.RemoveAll(s => s.Id == selected.Id); analysis = analysis.WithScenarios(scenarios); dirty = true; RefreshList();
        };
        remove.Click += (_, _) => { if (grid.CurrentRow != null) { grid.CurrentRow.Cells[2].Value = null; CommitRow(grid.CurrentRow.Index); } };
        grid.CellValidating += (_, e) => {
            if (loading || e.ColumnIndex != 2 || e.RowIndex < 0) return;
            try { ReadAdjustment(e.RowIndex, Convert.ToString(e.FormattedValue)); grid.Rows[e.RowIndex].ErrorText = ""; }
            catch (ArgumentException ex) { e.Cancel = true; grid.Rows[e.RowIndex].ErrorText = ex.Message; status.Text = ex.Message; }
        };
        grid.CellEndEdit += (_, e) => { if (!loading && e.RowIndex >= 0) CommitRow(e.RowIndex); };
        saveButton.Click += (_, _) => Guard(Save);
        runButton.Click += async (_, _) => await Run();
        close.Click += (_, _) => { if (cancellation != null) { cancellation.Cancel(); status.Text = "Cancelling; restoring and closing the temporary workbook…"; } else Close(); };
        FormClosing += (_, e) => {
            if (cancellation != null) { cancellation.Cancel(); e.Cancel = true; status.Text = "Cancelling; wait for cleanup, then close."; }
            else if (dirty && MessageBox.Show(this, "Discard unsaved scenario edits?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) e.Cancel = true;
        };
        RefreshList();
    }
    private ScenarioDefinition? Selected => list.SelectedItem as ScenarioDefinition;
    private void Guard(Action action)
    {
        try { action(); }
        catch (Exception ex) { System.Diagnostics.Trace.TraceError(ex.ToString()); MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); }
    }
    private void RefreshList(string? select = null)
    {
        var checkedIds = list.CheckedItems.Cast<ScenarioDefinition>().Select(s => s.Id).ToHashSet();
        var oldIds = list.Items.Cast<ScenarioDefinition>().Select(s => s.Id).ToHashSet();
        list.Items.Clear();
        foreach (var scenario in scenarios) list.Items.Add(scenario, scenario.Id == analysis.BaselineScenarioId || checkedIds.Contains(scenario.Id) || !oldIds.Contains(scenario.Id));
        baselineLabel.Text = "Baseline scenario: " + (scenarios.FirstOrDefault(s => s.Id == analysis.BaselineScenarioId)?.Name ?? "None — select a scenario and choose Set as Baseline");
        list.SelectedIndex = scenarios.Count == 0 ? -1 : Math.Max(0, scenarios.FindIndex(s => s.Id == select));
        ShowScenario();
    }
    private void ShowScenario()
    {
        if (loading) return;
        loading = true;
        try
        {
            grid.Rows.Clear(); grid.Enabled = Selected != null;
            if (Selected is not { } selected) return;
            name.Text = selected.Name;
            foreach (var a in assumptions)
            {
                string id = ScenarioSimulationService.Identity(a);
                var adjustment = selected.Adjustments.FirstOrDefault(x => x.AssumptionId.Equals(id, StringComparison.OrdinalIgnoreCase));
                int row = grid.Rows.Add($"{a.Name} ({a.SheetName}!{a.CellAddress})", a.Distribution.ToString(), adjustment?.Value.ToString("G", CultureInfo.CurrentCulture) ?? "");
                grid.Rows[row].Tag = id;
            }
            foreach (var missing in selected.Adjustments.Where(x => !assumptions.Any(a => ScenarioSimulationService.Identity(a).Equals(x.AssumptionId, StringComparison.OrdinalIgnoreCase))))
            {
                int row = grid.Rows.Add("Missing assumption — clear this adjustment", "Unavailable", missing.Value.ToString(CultureInfo.CurrentCulture));
                grid.Rows[row].Tag = missing.AssumptionId; grid.Rows[row].ErrorText = missing.AssumptionId;
            }
        }
        finally { loading = false; }
    }
    private ScenarioAdjustment? ReadAdjustment(int row, string? text)
    {
        if (string.IsNullOrWhiteSpace(text) || text == "No change") return null;
        if (!double.TryParse(text.Trim().TrimEnd('%'), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out double value))
            throw new ArgumentException("Enter a percentage such as 15 or -20, or clear the cell for No change.");
        var adjustment = new ScenarioAdjustment((string)grid.Rows[row].Tag!, ScenarioAdjustmentType.PercentageChange, value);
        ScenarioSimulationService.Validate(assumptions, new[] { new ScenarioDefinition("draft", "Draft", new[] { adjustment }) });
        return adjustment;
    }
    private void CommitRow(int row) => Guard(() => {
        if (Selected is not { } selected) return;
        var adjustment = ReadAdjustment(row, Convert.ToString(grid.Rows[row].Cells[2].Value));
        string id = (string)grid.Rows[row].Tag!;
        var edits = selected.Adjustments.Where(a => !a.AssumptionId.Equals(id, StringComparison.OrdinalIgnoreCase)).ToList();
        if (adjustment != null) edits.Add(adjustment);
        var updated = selected with { Adjustments = edits.ToArray() };
        scenarios[scenarios.FindIndex(s => s.Id == selected.Id)] = updated;
        loading = true;
        int index = list.SelectedIndex; bool check = list.GetItemChecked(index);
        // Avoid replacing ListBox items while a grid edit is committing: selection uses the current scenario by ID.
        list.Items[index] = updated; list.SetItemChecked(index, check);
        loading = false; dirty = true;
    });
    private void Save()
    {
        if (!ValidateChildren() || !grid.EndEdit()) throw new ArgumentException("Correct the highlighted percentage before saving.");
        string? id = Selected?.Id;
        analysis = save(question.Snapshot(analysis.WithScenarios(scenarios)));
        scenarios = analysis.Scenarios.ToList(); question.Reload(analysis); dirty = false; RefreshList(id);
        status.Text = "Scenario definitions saved in workbook. Save the workbook to retain them on disk.";
    }
    private async Task Run()
    {
        try
        {
            Save();
            var selected = list.CheckedItems.Cast<ScenarioDefinition>().ToArray();
            if (selected.Length == 0) throw new ArgumentException("Check at least one scenario to compare.");
            ScenarioSimulationService.Validate(assumptions, selected);
            var runAnalysis = analysis with { Scenarios = selected };
            ScenarioSimulationService.ValidateAnalysis(runAnalysis, forecasts);
            cancellation = new(); editor.Enabled = false; runButton.Enabled = false;
            foreach (Control control in runButton.Parent!.Controls) control.Enabled = control == close;
            close.Text = "Cancel run";
            var result = await run(runAnalysis, cancellation.Token, new Progress<string>(message => status.Text = message));
            status.Text = "Comparison complete. The source workbook's simulation cells were not changed.";
            using var comparison = new ScenarioComparisonForm(result); comparison.ShowDialog(this);
        }
        catch (OperationCanceledException) { status.Text = "Scenario Analysis cancelled; temporary resources cleaned up."; }
        catch (Exception ex) { System.Diagnostics.Trace.TraceError(ex.ToString()); status.Text = "Scenario Analysis did not finish."; MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        finally
        {
            cancellation?.Dispose(); cancellation = null; editor.Enabled = true;
            foreach (Control control in runButton.Parent!.Controls) control.Enabled = true;
            close.Text = "Close";
        }
    }
}

internal static class ScenarioUi
{
    public static TableLayoutPanel Table(int columns, int rows)
    {
        var table = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = columns, RowCount = rows, AutoSize = false };
        if (columns == 1) table.ColumnStyles.Add(new(SizeType.Percent, 100));
        if (rows == 1) table.RowStyles.Add(new(SizeType.Percent, 100));
        return table;
    }
    public static FlowLayoutPanel Flow() => new() { Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true };
    public static Button Button(string text) => new() { Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(6, 3, 6, 3) };
    public static DataGridView Grid(string name) => new()
    {
        Name = name, Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
        RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.CellSelect, BackgroundColor = SystemColors.Window
    };
}
