using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.Json;

namespace MonteCarlo.Excel;

// Only Interior is touched: values, formulas and all other formatting stay intact.
public static class ModelCellHighlight
{
    public const int AssumptionColor = 0xF7EBDD; // Excel OLE RGB: light blue
    public const int ForecastColor = 0xE2EFDA; // light green
    private const int MetadataColumn = 15;
    private const string Header = "OriginalFillV1";

    public sealed class Fill
    {
        public GradientFill? Gradient { get; set; }
        public int Pattern { get; set; }
        public double Color { get; set; }
        public int ColorIndex { get; set; }
        public int? Theme { get; set; }
        public double Tint { get; set; }
        public double PatternColor { get; set; }
        public int PatternColorIndex { get; set; }
        public int? PatternTheme { get; set; }
        public double PatternTint { get; set; }
    }

    public sealed class GradientFill
    {
        public double Degree { get; set; }
        public double Left { get; set; }
        public double Right { get; set; }
        public double Top { get; set; }
        public double Bottom { get; set; }
        public List<Stop> Stops { get; set; } = new();
    }
    public sealed class Stop
    {
        public double Position { get; set; }
        public double Color { get; set; }
        public int? Theme { get; set; }
        public double Tint { get; set; }
    }
    private static int? Theme(Func<object> read)
    {
        try { int value = Convert.ToInt32(read()); return value > 0 ? value : null; }
        catch (System.Runtime.InteropServices.COMException) { return null; }
    }

    private static Fill Capture(dynamic interior)
    {
        int pattern = Convert.ToInt32(interior.Pattern);
        if (pattern == 4000 || pattern == 4001)
        {
            dynamic gradient = interior.Gradient;
            var saved = new GradientFill();
            if (pattern == 4000) saved.Degree = Convert.ToDouble(gradient.Degree);
            else
            {
                saved.Left = Convert.ToDouble(gradient.RectangleLeft);
                saved.Right = Convert.ToDouble(gradient.RectangleRight);
                saved.Top = Convert.ToDouble(gradient.RectangleTop);
                saved.Bottom = Convert.ToDouble(gradient.RectangleBottom);
            }
            for (int i = 1; i <= (int)gradient.ColorStops.Count; i++)
            {
                dynamic stop = gradient.ColorStops.Item(i);
                saved.Stops.Add(new Stop { Position = Convert.ToDouble(stop.Position),
                    Color = Convert.ToDouble(stop.Color), Theme = Theme(() => stop.ThemeColor),
                    Tint = Convert.ToDouble(stop.TintAndShade) });
            }
            return new Fill { Pattern = pattern, Gradient = saved };
        }
        return new Fill {
            Pattern = pattern, Color = Convert.ToDouble(interior.Color),
            ColorIndex = Convert.ToInt32(interior.ColorIndex), Theme = Theme(() => interior.ThemeColor),
            Tint = Convert.ToDouble(interior.TintAndShade),
            PatternColor = Convert.ToDouble(interior.PatternColor),
            PatternColorIndex = Convert.ToInt32(interior.PatternColorIndex),
            PatternTheme = Theme(() => interior.PatternThemeColor),
            PatternTint = Convert.ToDouble(interior.PatternTintAndShade)
        };
    }

    private static void Restore(dynamic interior, Fill fill)
    {
        if (fill.Gradient != null)
        {
            interior.Pattern = fill.Pattern;
            dynamic gradient = interior.Gradient;
            if (fill.Pattern == 4000) gradient.Degree = fill.Gradient.Degree;
            else
            {
                gradient.RectangleLeft = fill.Gradient.Left;
                gradient.RectangleRight = fill.Gradient.Right;
                gradient.RectangleTop = fill.Gradient.Top;
                gradient.RectangleBottom = fill.Gradient.Bottom;
            }
            gradient.ColorStops.Clear();
            foreach (var saved in fill.Gradient.Stops)
            {
                dynamic stop = gradient.ColorStops.Add(saved.Position);
                if (saved.Theme.HasValue) stop.ThemeColor = saved.Theme.Value;
                else stop.Color = saved.Color;
                stop.TintAndShade = saved.Tint;
            }
            return;
        }
        if (fill.Theme.HasValue) interior.ThemeColor = fill.Theme.Value;
        else if (fill.ColorIndex < 0) interior.ColorIndex = fill.ColorIndex;
        else interior.Color = fill.Color;
        interior.TintAndShade = fill.Tint;
        if (fill.PatternTheme.HasValue) interior.PatternThemeColor = fill.PatternTheme.Value;
        else if (fill.PatternColorIndex < 0) interior.PatternColorIndex = fill.PatternColorIndex;
        else interior.PatternColor = fill.PatternColor;
        interior.PatternTintAndShade = fill.PatternTint;
        interior.Pattern = fill.Pattern;
    }

    private sealed record Item(string Type, string Sheet, string Address, string Link);
    private static IEnumerable<Item> Current() =>
        SimulationModel.Assumptions.Select(x => new Item("Assumption", x.SheetName, x.CellAddress, x.CellLink))
        .Concat(SimulationModel.Forecasts.Select(x => new Item("Forecast", x.SheetName, x.CellAddress, x.CellLink)));
    private static string Key(Item item) => string.IsNullOrEmpty(item.Link)
        ? item.Sheet + "!" + item.Address : item.Link;
    private static dynamic Cell(dynamic workbook, Item item)
    {
        dynamic cell;
        if (string.IsNullOrEmpty(item.Link)) cell = workbook.Worksheets[item.Sheet].Range[item.Address];
        else
        {
            if (!item.Link.StartsWith("_MC_Cell_", StringComparison.Ordinal)) throw new InvalidOperationException("Invalid model link.");
            cell = workbook.Names[item.Link].RefersToRange;
        }
        if (Convert.ToDouble(cell.Cells.CountLarge) != 1 ||
            !string.Equals((string)cell.Worksheet.Parent.Name, (string)workbook.Name, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Invalid model cell.");
        return cell;
    }

    // Called before rewriting the config: restore deleted items and carry surviving snapshots forward.
    public static Dictionary<string, string> BeforeSave(dynamic workbook, dynamic config, bool clear = false)
    {
        var retained = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var current = clear ? new List<Item>() : Current().ToList();
        int last = (int)config.UsedRange.Row + (int)config.UsedRange.Rows.Count - 1;
        if (Convert.ToString(config.Cells[1, MetadataColumn].Value2) != Header) return retained;
        for (int row = 2; row <= last; row++)
        {
            var item = new Item(Convert.ToString(config.Cells[row, 1].Value2) ?? "",
                Convert.ToString(config.Cells[row, 2].Value2) ?? "",
                Convert.ToString(config.Cells[row, 3].Value2) ?? "",
                Convert.ToString(config.Cells[row, 10].Value2) ?? "");
            if (item.Type != "Assumption" && item.Type != "Forecast") continue;
            string? json = Convert.ToString(config.Cells[row, MetadataColumn].Value2);
            if (string.IsNullOrEmpty(json)) continue;
            var survivor = current.FirstOrDefault(x => Key(x).Equals(Key(item), StringComparison.OrdinalIgnoreCase)
                || (x.Sheet.Equals(item.Sheet, StringComparison.OrdinalIgnoreCase) && x.Address.Equals(item.Address, StringComparison.OrdinalIgnoreCase)));
            if (survivor != null) retained[Key(survivor)] = json;
            else Try(() => Restore(Cell(workbook, item).Interior, JsonSerializer.Deserialize<Fill>(json)!));
        }
        return retained;
    }

    public static void ApplyRow(dynamic workbook, dynamic config, int row, string type, string sheet, string address,
        string link, Dictionary<string, string>? retained = null)
    {
        Try(() => {
            var item = new Item(type, sheet, address, link);
            dynamic cell = Cell(workbook, item);
            string? json = null;
            if (retained != null) retained.TryGetValue(Key(item), out json);
            else if (Convert.ToString(config.Cells[1, MetadataColumn].Value2) == Header)
                json = Convert.ToString(config.Cells[row, MetadataColumn].Value2);
            if (string.IsNullOrEmpty(json)) json = JsonSerializer.Serialize(Capture(cell.Interior));
            // Persist the snapshot before changing the fill, including during legacy-workbook migration.
            config.Cells[1, MetadataColumn].Value2 = Header;
            config.Cells[row, MetadataColumn].Value2 = json;
            cell.Interior.Pattern = 1;
            cell.Interior.Color = type == "Assumption" ? AssumptionColor : ForecastColor;
            cell.Interior.TintAndShade = 0;
        });
    }

    private static void Try(Action action)
    {
        try { action(); }
        catch (Exception ex) { Trace.TraceWarning("Model cell highlighting: {0}", ex.Message); }
    }
}
