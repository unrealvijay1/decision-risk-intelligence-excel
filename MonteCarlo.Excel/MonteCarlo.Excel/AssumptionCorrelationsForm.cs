using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public sealed class AssumptionCorrelationsForm : Form
{
    public const string HelpText = "Controls how strongly two assumptions tend to move together. The value defines dependence used during simulation; the observed Pearson correlation of simulated values may differ depending on their distributions.";
    private sealed record Choice(string Id, string Label);
    private readonly DataGridView grid = ScenarioUi.Grid("Correlations");
    private readonly IReadOnlyList<AssumptionDefinition> assumptions;
    private readonly Action<IReadOnlyList<AssumptionCorrelation>> save;
    public AssumptionCorrelationsForm(IReadOnlyList<AssumptionDefinition> assumptions,
        IReadOnlyList<AssumptionCorrelation> relationships, Action<IReadOnlyList<AssumptionCorrelation>> save)
    {
        this.assumptions = assumptions; this.save = save;
        Text = "Assumption Correlations"; ClientSize = new(780, 470); MinimumSize = new(650, 380);
        AutoScaleMode = AutoScaleMode.Dpi; StartPosition = FormStartPosition.CenterParent;
        var root = ScenarioUi.Table(1, 4); root.Padding = new Padding(12); Controls.Add(root);
        root.RowStyles.Add(new(SizeType.AutoSize)); root.RowStyles.Add(new(SizeType.Percent, 100));
        root.RowStyles.Add(new(SizeType.Absolute, 70)); root.RowStyles.Add(new(SizeType.AutoSize));
        root.Controls.Add(new Label { AutoSize = true, Text = "Define relationships between assumptions that tend to move together.\n+1 = move together    0 = no relationship    -1 = move in opposite directions" }, 0, 0);
        var choices = assumptions.Select(a => new Choice(ScenarioSimulationService.Identity(a),
            string.IsNullOrWhiteSpace(a.Name) ? $"{a.SheetName}!{a.CellAddress}" : $"{a.Name} ({a.SheetName}!{a.CellAddress})")).ToList();
        foreach (string id in relationships.SelectMany(r => new[] { r.AssumptionA, r.AssumptionB }).Distinct(StringComparer.OrdinalIgnoreCase))
            if (!choices.Any(c => c.Id.Equals(id, StringComparison.OrdinalIgnoreCase))) choices.Add(new(id, "Missing assumption: " + id));
        foreach (string title in new[] { "Assumption A", "Assumption B" })
            grid.Columns.Add(new DataGridViewComboBoxColumn { HeaderText = title, DataSource = choices.ToArray(), DisplayMember = "Label", ValueMember = "Id", FillWeight = 40 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Correlation", FillWeight = 20, ToolTipText = HelpText });
        foreach (var r in relationships) grid.Rows.Add(choices.First(c => c.Id.Equals(r.AssumptionA, StringComparison.OrdinalIgnoreCase)).Id,
            choices.First(c => c.Id.Equals(r.AssumptionB, StringComparison.OrdinalIgnoreCase)).Id, r.Correlation.ToString("R"));
        grid.DataError += (_, e) => { e.ThrowException = false; if (e.RowIndex >= 0) grid.Rows[e.RowIndex].ErrorText = "Select an existing assumption."; };
        root.Controls.Add(grid, 0, 1);
        root.Controls.Add(new Label { Name = "CorrelationHelp", Dock = DockStyle.Fill, Text = HelpText }, 0, 2);
        var footer = ScenarioUi.Flow(); root.Controls.Add(footer, 0, 3);
        var add = ScenarioUi.Button("Add"); var delete = ScenarioUi.Button("Delete");
        var accept = ScenarioUi.Button("Save"); var cancel = ScenarioUi.Button("Cancel"); cancel.DialogResult = DialogResult.Cancel;
        footer.Controls.AddRange(new Control[] { add, delete, accept, cancel }); CancelButton = cancel;
        add.Click += (_, _) => { int row = grid.Rows.Add(new object[] { null!, null!, "0" }); grid.CurrentCell = grid.Rows[row].Cells[0]; };
        delete.Click += (_, _) => { if (grid.CurrentRow != null) grid.Rows.Remove(grid.CurrentRow); };
        accept.Click += (_, _) => {
            try { var values = ReadRelationships(); save(values); DialogResult = DialogResult.OK; Close(); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        };
    }
    public IReadOnlyList<AssumptionCorrelation> ReadRelationships()
    {
        if (!grid.EndEdit()) throw new ArgumentException("Finish editing the correlation first.");
        var values = new List<AssumptionCorrelation>();
        foreach (DataGridViewRow row in grid.Rows)
        {
            if (!double.TryParse(Convert.ToString(row.Cells[2].Value), out double coefficient)) throw new ArgumentException("Enter a correlation between -1 and +1.");
            values.Add(new(Convert.ToString(row.Cells[0].Value) ?? "", Convert.ToString(row.Cells[1].Value) ?? "", coefficient));
        }
        _ = new GaussianDependence(assumptions.Select(ScenarioSimulationService.Identity).ToArray(), values);
        return values.Where(r => r.Correlation != 0).ToArray();
    }
}
