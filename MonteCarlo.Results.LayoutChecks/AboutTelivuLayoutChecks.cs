using System.Drawing;
using System.Windows.Forms;
using MonteCarlo.Excel;

internal static class AboutTelivuLayoutChecks
{
    public static void Run()
    {
        foreach (float scale in new[] { 1f, 1.25f, 1.5f })
        {
            using var form = new AboutTelivuForm();
            _ = form.Handle;
            form.Scale(new SizeF(scale, scale));
            form.Show(); form.PerformLayout(); Application.DoEvents();
            var controls = Descendants(form).ToArray();
            if (form.Text != "About Telivu") throw new Exception("Wrong About title.");
            var links = controls.OfType<LinkLabel>().ToArray();
            if (!links.Select(l => l.Text).SequenceEqual(new[] { "Website", "GitHub", "Report an Issue" }))
                throw new Exception("About links incomplete or unsupported donation link published.");
            if (controls.Any(c => c is TextBox || c.Text.Contains("Activate", StringComparison.OrdinalIgnoreCase)))
                throw new Exception("Activation UI remains.");
            foreach (var control in controls.Where(c => c is Label || c is Button))
            {
                var bounds = form.RectangleToClient(control.RectangleToScreen(control.ClientRectangle));
                if (!form.ClientRectangle.Contains(bounds)) throw new Exception("About control clipped: " + control.Text);
            }
            var close = controls.OfType<Button>().Single();
            if (close.DialogResult != DialogResult.OK || form.CancelButton != close)
                throw new Exception("About dismissal broken.");
            Directory.CreateDirectory("artifacts/about-layout");
            using var bitmap = new Bitmap(form.Width, form.Height);
            form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
            bitmap.Save($"artifacts/about-layout/about-{scale * 100:0}.png");
            form.Hide();
            Console.WriteLine($"ABOUT PASS {scale * 100:0}%: content, links, bounds and close action");
        }
    }
    private static IEnumerable<Control> Descendants(Control parent)
    {
        foreach (Control child in parent.Controls)
        {
            yield return child;
            foreach (var nested in Descendants(child)) yield return nested;
        }
    }
}
