using System.Reflection;
using MonteCarlo.Core.Tests;
using MonteCarlo.Excel;

internal static class ConstraintLayoutChecks
{
    public static void Run()
    {
        Directory.CreateDirectory("artifacts/constraint-layout");
        foreach (float scale in new[] { 1f, 1.25f, 1.5f })
        {
            var book = new OptimizationTests.Book(); var p = OptimizationTests.Problem();
            var linear = IterativeOptimizationTests.Linear(50, OptimizationConstraintOperator.Equal, ("x", 2));
            using var linearForm = new OptimizerConstraintForm(linear, p.Variables, [OptimizationTests.Forecast()], () => null, _ => throw new Exception("Linear constraint read Excel"));
            Prepare(linearForm, scale); Bounds(linearForm, Find(linearForm, "LinearExpression"));
            Click(Descendants(linearForm).OfType<Button>().Single(b => b.Text == "OK"));
            if (linearForm.Constraint?.Terms?.Single().Coefficient != 2 || linearForm.Constraint.EffectiveOperator != OptimizationConstraintOperator.Equal) throw new Exception("Linear constraint editor failed");
            int picks = 0;
            using var form = new OptimizerConstraintForm(null, p.Variables, [OptimizationTests.Forecast()], () => { picks++; return picks == 1 ? new("Model", "$C$2", 2, "Labor Used") : null; }, c => ConstraintCellSelection.Read(book, c.Sheet, c.Cell));
            Prepare(form, scale);
            foreach (string field in new[] { "ConstraintType", "ConstraintName", "ConstraintReference", "ConstraintOperator", "ConstraintValue", "SelectConstraintCell" }) Bounds(form, Find(form, field));
            Click((Button)Find(form, "SelectConstraintCell"));
            if (Find(form, "ConstraintName").Text != "Labor Used" || Find(form, "ConstraintReference").Text != "'Model'!$C$2") throw new Exception("Constraint selection/name suggestion missing.");
            Click((Button)Find(form, "SelectConstraintCell")); if (Find(form, "ConstraintReference").Text != "'Model'!$C$2") throw new Exception("Cancelled picker altered constraint.");
            Find(form, "ConstraintValue").Text = "2"; ((ComboBox)Find(form, "ConstraintOperator")).SelectedItem = "=";
            using (var bitmap = new Bitmap(form.Width, form.Height)) { form.DrawToBitmap(bitmap, new(Point.Empty, bitmap.Size)); bitmap.Save($"artifacts/constraint-layout/cell-{scale}.png"); }
            var size = form.ClientSize; form.ClientSize = new(size.Width, 280);
            using (var font = new Font(form.Font.FontFamily, 11))
            { var originalFont = form.Font; form.Font = font; Layout(form); foreach (var button in Descendants(form).OfType<Button>().Where(b => b.Text is "OK" or "Cancel")) Bounds(form, button); form.Font = originalFont; }
            form.ClientSize = size;
            Click(Descendants(form).OfType<Button>().Single(b => b.Text == "OK"));
            if (form.Constraint?.Kind != OptimizationConstraintKind.Cell || form.Constraint.Name != "Labor Used" || form.Constraint.EffectiveOperator != OptimizationConstraintOperator.Equal) throw new Exception("Cell/equality constraint not accepted.");
            using var typed = new OptimizerConstraintForm(null, [], [], () => null, c => ConstraintCellSelection.Read(book, c.Sheet, c.Cell)); Prepare(typed, scale);
            Find(typed, "ConstraintReference").Text = "Model!C2"; Find(typed, "ConstraintValue").Text = "2";
            Click(Descendants(typed).OfType<Button>().Single(b => b.Text == "OK")); if (typed.Constraint?.Cell != "$C$2") throw new Exception("Typed reference needs a Forecast/Variable registration.");
            var old = CellConstraintTests.Cell("C2", 2, OptimizationConstraintOperator.Equal, "My Labor") with { CellLink = "tracked" };
            using var edited = new OptimizerConstraintForm(old, p.Variables, [OptimizationTests.Forecast()], () => new("Renamed", "$C$2", 2, "Other"), _ => new("Renamed", "$C$2", 2, "Suggestion")); Prepare(edited, scale);
            Click((Button)Find(edited, "SelectConstraintCell")); Click(Descendants(edited).OfType<Button>().Single(b => b.Text == "OK"));
            if (edited.Constraint?.CellLink != "tracked" || edited.Constraint.Name != "My Labor") throw new Exception("Edit lost tracked link or entered name.");
            foreach (var kind in new[] { OptimizationConstraintKind.DecisionVariable, OptimizationConstraintKind.Mean, OptimizationConstraintKind.Median, OptimizationConstraintKind.Probability })
            {
                var legacy = new OptimizationConstraint(kind, kind == OptimizationConstraintKind.DecisionVariable ? "x" : "forecast", MonteCarlo.Core.TargetDirection.AtOrBelow, kind == OptimizationConstraintKind.Probability ? 80 : 30);
                using var dialog = new OptimizerConstraintForm(legacy, p.Variables, [OptimizationTests.Forecast()], () => null, _ => throw new Exception("Legacy constraint read a raw cell.")); Prepare(dialog, scale);
                Bounds(dialog, Find(dialog, "ConstraintSubject")); Bounds(dialog, Find(dialog, "ConstraintOperator"));
                Click(Descendants(dialog).OfType<Button>().Single(b => b.Text == "OK"));
                if (dialog.Constraint?.Kind != kind || dialog.Constraint.SubjectId != legacy.SubjectId || dialog.Constraint.EffectiveOperator != OptimizationConstraintOperator.LessThanOrEqual) throw new Exception("Legacy constraint editing changed behavior.");
            }
            Console.WriteLine($"CONSTRAINT UI PASS scale={scale}: typed/picked cells, cancellation, equality, tracked edit, legacy types and footer bounds");
        }
    }
    private static Control Find(Control root, string name) => Descendants(root).Single(c => c.Name == name);
    private static void Click(Button button) => typeof(Button).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(button, [EventArgs.Empty]);
    private static IEnumerable<Control> Descendants(Control root) { foreach (Control child in root.Controls) { yield return child; foreach (var c in Descendants(child)) yield return c; } }
    private static void Layout(Form form) { for (int i = 0; i < 4; i++) { form.PerformLayout(); foreach (var c in Descendants(form)) c.PerformLayout(); } }
    private static void Prepare(Form form, float scale) { _ = form.Handle; foreach (var c in Descendants(form)) _ = c.Handle; form.Scale(new SizeF(scale, scale)); Layout(form); }
    private static void Bounds(Form form, Control control)
    {
        var point = form.PointToClient(control.PointToScreen(Point.Empty));
        if (control.Width < 30 || control.Height < 18 || point.X < 0 || point.Y < 0 || point.X + control.Width > form.ClientSize.Width || point.Y + control.Height > form.ClientSize.Height)
            throw new Exception($"Constraint control clipped: {control.Name}/{control.Text}, {point}, {control.Size}");
    }
}
