using System.Globalization;
namespace MonteCarlo.Excel;

public sealed class DecisionVariableForm : Form
{
    private readonly TextBox name = new() { Dock = DockStyle.Fill, Name = "VariableName" }, cell = new() { Dock = DockStyle.Fill, ReadOnly = true, Name = "SelectedCell" },
        minimum = new() { Dock = DockStyle.Fill }, maximum = new() { Dock = DockStyle.Fill }, step = new() { Dock = DockStyle.Fill };
    private readonly Label current = new() { AutoSize = true, Text = "Select a numeric input", Name = "CurrentValue" };
    private DecisionVariableSelection? selection;
    private DecisionVariableSelection? originalSelection;
    public DecisionVariable? Variable { get; private set; }
    public DecisionVariableForm(DecisionVariable? existing, Func<DecisionVariableSelection?> pick,
        Func<DecisionVariable, DecisionVariableSelection> read, Action<DecisionVariable> validate)
    {
        Text = "Decision Variable"; AutoScaleMode = AutoScaleMode.Dpi; ClientSize = new(510, 360); MinimumSize = new(470, 350); StartPosition = FormStartPosition.CenterParent;
        var layout = ScenarioUi.Table(1, 2); layout.RowStyles.Add(new(SizeType.Percent, 100)); layout.RowStyles.Add(new(SizeType.AutoSize)); Controls.Add(layout);
        var body = new Panel { Dock = DockStyle.Fill, AutoScroll = true }; layout.Controls.Add(body, 0, 0);
        var root = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new(12), ColumnCount = 2, RowCount = 7 };
        root.ColumnStyles.Add(new(SizeType.AutoSize)); root.ColumnStyles.Add(new(SizeType.Percent, 100)); body.Controls.Add(root);
        for (int i = 0; i < 7; i++) root.RowStyles.Add(new(SizeType.AutoSize));
        void Row(int row, string label, Control control) { root.Controls.Add(new Label { Text = label, AutoSize = true, Margin = new(3, 7, 12, 7) }, 0, row); root.Controls.Add(control, 1, row); }
        name.Text = existing?.Name ?? ""; minimum.Text = (existing?.Minimum ?? 0).ToString("G17"); maximum.Text = (existing?.Maximum ?? 100).ToString("G17"); step.Text = (existing?.Step ?? 1).ToString("G17");
        Row(0, "Variable Name", name);
        var reference = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2, RowCount = 1, Margin = new(0) };
        reference.ColumnStyles.Add(new(SizeType.Percent, 100)); reference.ColumnStyles.Add(new(SizeType.AutoSize));
        var select = ScenarioUi.Button("Select"); reference.Controls.Add(cell, 0, 0); reference.Controls.Add(select, 1, 0); Row(1, "Cell", reference); Row(2, "Current Value", current);
        root.Controls.Add(new Label { Text = "Search Range", AutoSize = true, Margin = new(3, 12, 3, 6) }, 0, 3);
        Row(4, "Minimum", minimum); Row(5, "Maximum", maximum); Row(6, "Step", step);
        var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.RightToLeft, Padding = new(9, 3, 9, 9) };
        var ok = ScenarioUi.Button("OK"); var cancel = ScenarioUi.Button("Cancel"); cancel.DialogResult = DialogResult.Cancel;
        footer.Controls.AddRange([ok, cancel]); layout.Controls.Add(footer, 0, 1); AcceptButton = ok; CancelButton = cancel;
        void Set(DecisionVariableSelection selected) { selection = selected; cell.Text = selected.Display; current.Text = selected.CurrentValue.ToString("G17"); name.Text = DecisionVariableSelection.SuggestName(name.Text, selected.SuggestedName); }
        if (existing != null)
        {
            cell.Text = existing.Sheet + "!" + existing.Cell;
            try { originalSelection = read(existing); Set(originalSelection); } catch (Exception ex) { current.Text = "Reference unavailable — select again"; current.Tag = ex.Message; }
        }
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
                if (selection == null) throw new ArgumentException("Select one editable numeric input cell before clicking OK.");
                double Number(TextBox box, string label) => double.TryParse(box.Text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out var value) && double.IsFinite(value) ? value : throw new ArgumentException($"{label}: enter a finite numeric value.");
                bool same = originalSelection != null && originalSelection.Sheet.Equals(selection.Sheet, StringComparison.OrdinalIgnoreCase) && originalSelection.Cell == selection.Cell;
                var draft = new DecisionVariable(existing?.Id ?? Guid.NewGuid().ToString("N"), name.Text.Trim(), selection.Sheet, selection.Cell,
                    Number(minimum, "Minimum"), Number(maximum, "Maximum"), Number(step, "Step"), same ? existing!.CellLink : "");
                validate(draft); Variable = draft; DialogResult = DialogResult.OK; Close();
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        };
    }
}
