using MonteCarlo.Core;
using MonteCarlo.Excel;
using System.Reflection;

internal static class DistributionLayoutChecks
{
    public static void Run()
    {
        Directory.CreateDirectory("artifacts/distribution-layout");
        foreach (float scale in new[] { 1f, 1.25f, 1.5f })
        foreach (DistributionKind kind in Enum.GetValues<DistributionKind>())
        {
            var p = DistributionCatalog.Defaults(kind);
            var input = new AssumptionDefinition { Name = "Example", Distribution = (DistributionType)kind,
                Parameter1 = p[0], Parameter2 = p[1], Parameter3 = p[2], Parameter4 = p[3],
                ProbabilityTable = kind == DistributionKind.Discrete ? new([new(-10, .2), new(2.5, .5), new(30, .3)]) : null };
            using var form = new AssumptionForm("A1", input);
            form.Show(); form.Scale(new SizeF(scale, scale)); form.PerformLayout();
            var selector = Field<ComboBox>(form, "cmbDistribution");
            if (selector.Items.Count != 14 || selector.Text != DistributionCatalog.Name(kind)) throw new Exception("Distribution selector incorrect.");
            var parameters = DistributionCatalog.Parameters(kind);
            for (int i = 1; i <= 4; i++)
            {
                var label = Field<Label>(form, "lblParameter" + i); var field = Field<TextBox>(form, "txtParameter" + i);
                if (field.Visible != (i <= parameters.Length) || label.Visible != field.Visible) throw new Exception("Parameter visibility incorrect.");
                if (field.Visible && label.Text != parameters[i - 1]) throw new Exception("Parameter label incorrect.");
            }
            var preview = Field<DistributionPreviewResult>(form, "preview");
            if (!preview.IsValid || preview.IsDiscrete != DistributionCatalog.IsDiscrete(kind)) throw new Exception("Preview invalid: " + preview.ValidationMessage);
            foreach (var control in Descendants(form).Where(c => c.Visible && c is Button or TextBox or ComboBox or Label or DataGridView))
            {
                var location = form.PointToClient(control.PointToScreen(Point.Empty));
                if (location.X < 0 || location.Y < 0 || location.X + control.Width > form.ClientSize.Width + 1 || location.Y + control.Height > form.ClientSize.Height + 1)
                    throw new Exception($"Clipped {kind} control at {scale}: {control.Text} {location} {control.Size} client {form.ClientSize}");
            }
            using (var bitmap = new Bitmap(form.Width, form.Height)) { form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); if (kind is DistributionKind.Discrete or DistributionKind.TruncatedNormal) bitmap.Save($"artifacts/distribution-layout/{kind}-{scale}.png"); }
            if (kind == DistributionKind.Discrete)
            {
                var grid = Field<DataGridView>(form, "probabilityGrid");
                Click(Descendants(form).OfType<Button>().Single(b => b.Text == "Add Row"));
                if (grid.Rows.Count != 4 || Field<DistributionPreviewResult>(form, "preview").IsValid) throw new Exception("Add-row validation missing.");
                grid.CurrentCell = grid.Rows[3].Cells[0]; Click(Descendants(form).OfType<Button>().Single(b => b.Text == "Delete Row"));
                grid.Rows[0].Cells[1].Value = .1;
                if (Field<DistributionPreviewResult>(form, "preview").IsValid || !Field<Label>(form, "probabilityTotal").Text.Contains("0.9")) throw new Exception("Probability total validation missing.");
                grid.Rows[0].Cells[1].Value = .2;
                if (!Field<DistributionPreviewResult>(form, "preview").IsValid) throw new Exception("Table correction failed.");
            }
            Click(Descendants(form).OfType<Button>().Single(b => b.Text == "OK"));
            if (form.DialogResult != DialogResult.OK || form.Distribution != (DistributionType)kind || form.AssumptionName != "Example") throw new Exception("Save failed.");
            if (kind == DistributionKind.Discrete && !input.ProbabilityTable!.Outcomes.SequenceEqual(form.ProbabilityTable!.Outcomes)) throw new Exception("Table edit round trip failed.");
            Console.WriteLine($"DISTRIBUTION UI PASS {kind} scale {scale}: fields, preview, validation, edit/save and bounds");
        }
    }
    private static T Field<T>(object obj, string name) => (T)obj.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(obj)!;
    private static void Click(Button b) => typeof(Button).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(b, [EventArgs.Empty]);
    private static IEnumerable<Control> Descendants(Control root) { foreach (Control child in root.Controls) { yield return child; foreach (var c in Descendants(child)) yield return c; } }
}

