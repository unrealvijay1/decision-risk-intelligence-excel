using MonteCarlo.Core;
using MonteCarlo.Excel;
using System.Reflection;

internal static class CorrelationLayoutChecks
{
    public static void Run()
    {
        foreach (float scale in new[] { 1f, 1.25f, 1.5f })
        {
            var inputs = new[] { new AssumptionDefinition { Name = "Scope", CellLink = "a", SheetName = "Model", CellAddress = "A1" },
                new AssumptionDefinition { Name = "Testing Effort", CellLink = "b", SheetName = "Model", CellAddress = "A2" } };
            using var form = new AssumptionCorrelationsForm(inputs, [new("a", "b", .8)], _ => { });
            _ = form.Handle;
            foreach (var control in Descendants(form)) _ = control.Handle;
            form.Scale(new SizeF(scale, scale));
            for (int i = 0; i < 4; i++) { form.PerformLayout(); foreach (var control in Descendants(form)) control.PerformLayout(); }
            var grid = Descendants(form).OfType<DataGridView>().Single();
            if (grid.Columns.Count != 3 || grid.Rows.Count != 1) throw new Exception("Correlation table missing.");
            foreach (var control in Descendants(form).Where(c => c is Button || c == grid || c.Name == "CorrelationHelp"))
            {
                var p = form.PointToClient(control.PointToScreen(Point.Empty));
                if (p.X < 0 || p.Y < 0 || p.X + control.Width > form.ClientSize.Width || p.Y + control.Height > form.ClientSize.Height)
                    throw new Exception("Clipped correlation control: " + control.Text);
            }
            grid.Rows[0].Cells[2].Value = "-0.4";
            if (form.ReadRelationships().Single().Correlation != -.4) throw new Exception("Correlation edit failed.");
            Click(Descendants(form).OfType<Button>().Single(b => b.Text == "Delete"));
            if (form.ReadRelationships().Count != 0) throw new Exception("Correlation delete failed.");
            Click(Descendants(form).OfType<Button>().Single(b => b.Text == "Add"));
            grid.Rows[0].Cells[0].Value = "a"; grid.Rows[0].Cells[1].Value = "b"; grid.Rows[0].Cells[2].Value = ".5";
            if (form.ReadRelationships().Single().Correlation != .5) throw new Exception("Correlation add failed.");
            Console.WriteLine($"CORRELATION PASS scale {scale}: add/edit/delete, table and help/footer bounds");
        }
    }
    private static void Click(Button b) => typeof(Button).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(b, [EventArgs.Empty]);
    private static IEnumerable<Control> Descendants(Control root) { foreach (Control child in root.Controls) { yield return child; foreach (var c in Descendants(child)) yield return c; } }
}
