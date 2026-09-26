using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using MonteCarlo.Excel;

// Test-only Excel-DNA boundary. Production persistence and simulation adapters run unchanged.
namespace ExcelDna.Integration
{
    public static class ExcelDnaUtil
    {
        public static object Application { get; set; } = null!;
    }
}

namespace MonteCarlo.Core.Tests
{
    public sealed class PersistenceApplication
    {
        public PersistenceWorkbook? ActiveWorkbook { get; set; }
        public bool ScreenUpdating { get; set; } = true;
        public bool EnableEvents { get; set; } = true;
        public bool DisplayAlerts { get; set; } = true;
        public object? StatusBar { get; set; } = false;
        public PersistenceFunctions WorksheetFunction { get; } = new();
        public Action? OnCalculate { get; set; }
        public int CalculationCount { get; private set; }
        public void Calculate() { CalculationCount++; OnCalculate?.Invoke(); }
    }
    public sealed class PersistenceFunctions
    {
        public bool IsError(PersistenceCell cell) => cell.Value2 is ExcelCellError;
    }
    public sealed class PersistenceWorkbook
    {
        public string Name { get; set; } = "Model.xlsx";
        public PersistenceSheets Worksheets { get; }
        public PersistenceNames Names { get; }
        public PersistenceWorkbook() { Worksheets = new(this); Names = new(this); }

        // Represents disk persistence: fresh objects, no SimulationModel state retained.
        public PersistenceWorkbook Reopen()
        {
            var copy = new PersistenceWorkbook { Name = Name };
            foreach (var sheet in Worksheets.Items)
            {
                var next = copy.Worksheets.Add(); next.Name = sheet.Name; next.Visible = sheet.Visible;
                foreach (var pair in sheet.Cells.Items)
                {
                    var cell = next.Cells[pair.Key.Row, pair.Key.Column];
                    cell.Value2 = pair.Value.Value2; cell.Location = pair.Value.Location;
                    cell.Deleted = pair.Value.Deleted; cell.NumberFormat = pair.Value.NumberFormat;
                }
            }
            foreach (var pair in Names.Items)
            {
                var old = pair.Value.Cell;
                var cell = copy.Worksheets[old.Worksheet.Name].Cells[old.Row, old.Column];
                copy.Names.Items.Add(pair.Key, new(cell, pair.Value.Visible));
            }
            return copy;
        }
    }
    public sealed class PersistenceSheets(PersistenceWorkbook workbook)
    {
        public List<PersistenceSheet> Items { get; } = new();
        public int Count => Items.Count;
        public PersistenceSheet this[int index] => Items[index - 1];
        public PersistenceSheet this[string name] => Items.FirstOrDefault(x => x.Name == name) ?? throw new COMException("Missing sheet");
        public PersistenceSheet Add() { var sheet = new PersistenceSheet(workbook); Items.Add(sheet); return sheet; }
    }
    public sealed class PersistenceSheet
    {
        public string Name { get; set; } = "Sheet1";
        public PersistenceWorkbook Parent { get; }
        public PersistenceCells Cells { get; }
        public PersistenceRanges Range { get; }
        public int Visible { get; set; } = -1;
        public bool ProtectContents => false;
        public PersistenceUsedRange UsedRange => new(Cells.Items.Where(x => x.Value.Value2 != null).Select(x => x.Key.Row).DefaultIfEmpty(1).Max());
        public PersistenceSheet(PersistenceWorkbook parent) { Parent = parent; Cells = new(this); Range = new(this); }
        public void Delete() { foreach (var cell in Cells.Items.Values) cell.Deleted = true; Parent.Worksheets.Items.Remove(this); }
    }
    public sealed class PersistenceUsedRange(int count)
    {
        public int Row => 1;
        public PersistenceCount Rows => new(count);
    }
    public sealed class PersistenceCount(int count)
    {
        public int Count => count;
        public double CountLarge => count;
    }
    public sealed class PersistenceCells(PersistenceSheet sheet)
    {
        public Dictionary<(int Row, int Column), PersistenceCell> Items { get; } = new();
        public PersistenceCell this[int row, int column]
        {
            get
            {
                if (!Items.TryGetValue((row, column), out var cell))
                    Items[(row, column)] = cell = new(sheet, row, column);
                return cell;
            }
        }
        public void Clear() => Items.Clear();
    }
    public sealed class PersistenceRanges(PersistenceSheet sheet)
    {
        public PersistenceCell this[string address]
        {
            get
            {
                if (address is "A:D" or "I:N") return new(sheet, 1, 1);
                var moved = sheet.Cells.Items.Values.FirstOrDefault(x => x.Location == address && !x.Deleted);
                if (moved != null) return moved;
                var match = Regex.Match(address, @"^\$?([A-Z]+)\$?([1-9][0-9]*)$");
                if (!match.Success) throw new COMException("Invalid range");
                int column = 0;
                foreach (char c in match.Groups[1].Value) column = column * 26 + c - 'A' + 1;
                return sheet.Cells[int.Parse(match.Groups[2].Value), column];
            }
        }
    }
    public sealed class PersistenceCell
    {
        public PersistenceSheet Worksheet { get; }
        public int Row { get; }
        public int Column { get; }
        public string Location { get; set; }
        public bool Deleted { get; set; }
        public object? Value2 { get; set; }
        public object? Formula { get; set; }
        public object? Formula2 { get; set; }
        public bool HasFormula => false;
        public bool HasArray => false;
        public bool HasSpill => false;
        public bool Locked => false;
        public bool MergeCells => false;
        public string NumberFormat { get; set; } = "General";
        public string Text => Value2 is ExcelCellError error ? error.Name : Convert.ToString(Value2) ?? "";
        public PersistenceCount Cells => new(1);
        public PersistenceCount Areas => new(1);
        public PersistenceAddress Address => new(this);
        public PersistenceCell(PersistenceSheet sheet, int row, int column)
        {
            Worksheet = sheet; Row = row; Column = column;
            string letters = "";
            for (int c = column; c > 0; c = (c - 1) / 26) letters = (char)('A' + (c - 1) % 26) + letters;
            Location = letters + row;
        }
    }
    public sealed class PersistenceAddress(PersistenceCell cell)
    {
        public string this[bool rowAbsolute, bool columnAbsolute] => cell.Location;
    }
    public sealed class PersistenceName(PersistenceCell cell, bool visible)
    {
        public PersistenceCell Cell => cell;
        public bool Visible => visible;
        public PersistenceCell RefersToRange => cell.Deleted ? throw new COMException("#REF!") : cell;
    }
    public sealed class PersistenceNames(PersistenceWorkbook workbook)
    {
        public Dictionary<string, PersistenceName> Items { get; } = new();
        public PersistenceName this[string name] => Items.TryGetValue(name, out var value) ? value : throw new COMException("Missing name");
        public void Add(string Name, string RefersTo, bool Visible)
        {
            int separator = RefersTo.LastIndexOf("'!", StringComparison.Ordinal);
            string sheet = RefersTo[2..separator].Replace("''", "'");
            Items.Add(Name, new(workbook.Worksheets[sheet].Range[RefersTo[(separator + 2)..]], Visible));
        }
    }
}
