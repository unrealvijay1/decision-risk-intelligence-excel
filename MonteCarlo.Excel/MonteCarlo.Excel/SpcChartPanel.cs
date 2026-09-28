using System.Drawing.Drawing2D;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public sealed class SpcChartPanel : Panel
{
    public SpcResult Result { get; }
    public SpcChart Chart { get; }
    public RectangleF PlotBounds { get; private set; }
    public SpcChartPanel(SpcResult result, SpcChart chart)
    {
        Result = result; Chart = chart; Dock = DockStyle.Fill; BackColor = Color.White;
        DoubleBuffered = true; MinimumSize = new Size(200, 100);
        AccessibleName = chart + " chart";
    }
    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
        float scale = Font.Size / 9f;
        var plot = new RectangleF(85 * scale, 70 * scale, Width - 112 * scale, Height - 116 * scale);
        PlotBounds = plot;
        if (plot.Width < 30 || plot.Height < 30) return;
        var b = Result.Baseline;
        bool individual = Chart == SpcChart.Individuals;
        double center = individual ? b.Mean : b.AverageMovingRange;
        double lower = individual ? b.IndividualsLower : 0;
        double upper = individual ? b.IndividualsUpper : b.MovingRangeUpper;
        double? target = individual ? Result.Configuration.Target?.Value : null;
        double[] values = individual ? Result.Observations.ToArray() : Result.MovingRanges.Skip(1).Select(v => v!.Value).ToArray();
        // Scale first, avoiding overflow for finite observations with very large magnitudes.
        double magnitude = values.Append(lower).Append(upper).Append(target ?? center).Max(v => Math.Abs(v));
        if (magnitude == 0) magnitude = 1;
        double min = Math.Min(values.Min() / magnitude, Math.Min(lower / magnitude, (target ?? lower) / magnitude));
        double max = Math.Max(values.Max() / magnitude, Math.Max(upper / magnitude, (target ?? upper) / magnitude));
        double padding = Math.Max((max - min) * .08, 1e-12); min -= padding; max += padding;
        float X(int index) => plot.Left + (index - 1f) / (Result.Observations.Count - 1) * plot.Width;
        float Y(double value) => plot.Bottom - (float)((value / magnitude - min) / (max - min)) * plot.Height;
        using var baselineFill = new SolidBrush(Color.FromArgb(18, 30, 100, 160));
        g.FillRectangle(baselineFill, X(b.Start), plot.Top, X(b.End) - X(b.Start), plot.Height);
        using var axis = new Pen(Color.LightGray);
        g.DrawRectangle(axis, plot.X, plot.Y, plot.Width, plot.Height);
        g.DrawString(individual ? "Individuals — " + Result.Configuration.MeasurementUnit : "Moving range", Font, Brushes.Black, 8, 3);
        g.DrawString("Blue: center   Red dashed: limits   Orange: business target   Shaded: baseline", Font, Brushes.DimGray, 8, 22 * scale);
        g.DrawString($"CL {center:G5}    LCL {lower:G5}    UCL {upper:G5}" + (target.HasValue ? $"    Target {target.Value:G5}" : ""), Font, Brushes.DimGray, 8, 42 * scale);
        void Line(double value, Color color, DashStyle dash, string label)
        {
            using var pen = new Pen(color, 1.4f) { DashStyle = dash };
            float y = Y(value); g.DrawLine(pen, plot.Left, y, plot.Right, y);
        }
        Line(center, Color.SteelBlue, DashStyle.Solid, "CL");
        Line(lower, Color.Firebrick, DashStyle.Dash, "LCL");
        Line(upper, Color.Firebrick, DashStyle.Dash, "UCL");
        if (target.HasValue) Line(target.Value, Color.DarkOrange, DashStyle.DashDot, "Target");
        for (int tick = 0; tick <= 4; tick++)
        {
            double normalized = min + (max - min) * tick / 4;
            double tickValue = normalized * magnitude;
            if (double.IsFinite(tickValue))
                g.DrawString(tickValue.ToString("G4"), Font, Brushes.DimGray, 2, plot.Bottom - tick * plot.Height / 4 - Font.Height / 2f);
        }
        using var boundary = new Pen(Color.SlateGray) { DashStyle = DashStyle.Dot };
        foreach (int index in new[] { b.Start, b.End }) g.DrawLine(boundary, X(index), plot.Top, X(index), plot.Bottom);
        var signaled = Result.Signals.Where(s => s.Chart == Chart).SelectMany(s => Chart == SpcChart.MovingRange
            ? new[] { s.ObservationIndex } : Enumerable.Range(s.StartIndex, s.EndIndex - s.StartIndex + 1)).ToHashSet();
        using var dataPen = new Pen(Color.FromArgb(35, 80, 115), 1.4f);
        PointF? previous = null;
        for (int i = individual ? 0 : 1; i < Result.Observations.Count; i++)
        {
            double value = individual ? Result.Observations[i] : Result.MovingRanges[i]!.Value;
            var point = new PointF(X(i + 1), Y(value));
            if (previous.HasValue) g.DrawLine(dataPen, previous.Value, point);
            if (signaled.Contains(i + 1) || Result.Observations.Count <= 300)
                g.FillEllipse(signaled.Contains(i + 1) ? Brushes.Firebrick : Brushes.SteelBlue, point.X - 3, point.Y - 3, 6, 6);
            previous = point;
        }
        int step = Math.Max(1, (int)Math.Ceiling(Result.Observations.Count / Math.Max(2, plot.Width / (110 * scale))));
        using var format = new StringFormat { Alignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap };
        for (int i = 0; i < Result.Observations.Count; i += step)
            g.DrawString($"{i + 1}: {Result.Labels[i]}", Font, Brushes.DimGray,
                new RectangleF(Math.Clamp(X(i + 1) - 55 * scale, 0, Width - 110 * scale), plot.Bottom + 6, 110 * scale, 30 * scale), format);
    }
}
