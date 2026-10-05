using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

namespace MonteCarlo.Excel;

/// <summary>Product information only; no account, activation or persistent state.</summary>
public sealed class AboutTelivuForm : Form
{
    public AboutTelivuForm()
    {
        Text = "About Telivu";
        Font = new System.Drawing.Font("Segoe UI", 10F);
        AutoScaleMode = AutoScaleMode.Dpi;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new System.Drawing.Size(520, 430);

        var content = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, Padding = new Padding(24), ColumnCount = 1,
            RowCount = 8, AutoSize = false
        };
        content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        for (int i = 0; i < 7; i++) content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        content.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        Label Line(string text, float size = 10F) => new()
        {
            Text = text, AutoSize = true, MaximumSize = new System.Drawing.Size(465, 0),
            Font = new System.Drawing.Font("Segoe UI", size), Margin = new Padding(0, 0, 0, 12)
        };
        content.Controls.Add(Line("TELIVU", 22F));
        content.Controls.Add(Line("Decision Intelligence for Excel", 12F));
        string version = typeof(AboutTelivuForm).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
            ?? typeof(AboutTelivuForm).Assembly.GetName().Version?.ToString() ?? "";
        content.Controls.Add(Line("Version: " + version));
        content.Controls.Add(Line("Free • No subscription • No activation"));
        content.Controls.Add(Line("Simulation, optimization, scenario analysis, sensitivity and statistical process control directly in Excel."));
        content.Controls.Add(Line("Open-source license: awaiting owner approval."));
        var links = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = true };
        foreach (var (label, url) in new[]
        {
            ("Website", "https://unrealvijay1.github.io/telivu/"),
            ("GitHub", "https://github.com/unrealvijay1/decision-risk-intelligence-excel"),
            ("Report an Issue", "https://github.com/unrealvijay1/decision-risk-intelligence-excel/issues")
        })
        {
            var link = new LinkLabel { Text = label, AutoSize = true, Margin = new Padding(0, 0, 18, 12) };
            link.LinkClicked += (_, _) =>
            {
                try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
                catch (Exception ex)
                {
                    Trace.TraceError(ex.ToString());
                    MessageBox.Show(this, "Your browser could not open the link.\n" + url,
                        "Telivu", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            links.Controls.Add(link);
        }
        content.Controls.Add(links);
        var close = new Button { Text = "Close", DialogResult = DialogResult.OK, AutoSize = true, Anchor = AnchorStyles.Right | AnchorStyles.Bottom };
        content.Controls.Add(close);
        AcceptButton = CancelButton = close;
        Controls.Add(content);
    }
}
