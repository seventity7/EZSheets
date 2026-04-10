using System.Globalization;
using System.Text.RegularExpressions;
using EZSheets.Models;

namespace EZSheets.Services;

public static partial class AdvancedFormulaEngine
{
    public static bool TryEvaluate(string expression, SheetDocument document, int activeTabIndex, out string result)
    {
        result = string.Empty;
        if (string.IsNullOrWhiteSpace(expression) || !expression.StartsWith("=", StringComparison.Ordinal))
        {
            return false;
        }

        try
        {
            var parser = new ExpressionParser(expression[1..], document, activeTabIndex);
            var value = parser.ParseExpressionValue();
            result = FormatNumber(value);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string FormatNumber(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            return "0";
        }

        if (Math.Abs(value % 1d) < 0.0000001d)
        {
            return Math.Round(value).ToString(CultureInfo.InvariantCulture);
        }

        return value.ToString("0.########", CultureInfo.InvariantCulture);
    }

    private static double ResolveCellNumber(SheetDocument document, int activeTabIndex, string token)
    {
        var targetTabIndex = activeTabIndex;
        var cellToken = token.Trim();
        var bangIndex = cellToken.LastIndexOf('!');
        if (bangIndex > 0)
        {
            var tabName = cellToken[..bangIndex].Trim();
            targetTabIndex = document.Tabs.FindIndex(x => string.Equals(x.Name, tabName, StringComparison.OrdinalIgnoreCase));
            if (targetTabIndex < 0)
            {
                targetTabIndex = activeTabIndex;
            }

            cellToken = cellToken[(bangIndex + 1)..].Trim();
        }

        if (targetTabIndex < 0 || targetTabIndex >= document.Tabs.Count)
        {
            return 0d;
        }

        var internalKey = TryParseCell(cellToken, out var row, out var column) ? ToInternalCellKey(row, column) : cellToken;
        if (!document.Tabs[targetTabIndex].TryGetCell(internalKey, out var cell) || cell is null)
        {
            return 0d;
        }

        var raw = (cell.Value ?? string.Empty).Trim();
        if (raw.StartsWith("=", StringComparison.Ordinal))
        {
            if (TryEvaluate(raw, document, targetTabIndex, out var nested)
                && double.TryParse((nested ?? string.Empty).Replace(",", string.Empty, StringComparison.Ordinal).Replace(" ", string.Empty, StringComparison.Ordinal), NumberStyles.Any, CultureInfo.InvariantCulture, out var nestedValue))
            {
                return nestedValue;
            }

            return 0d;
        }

        var numericText = raw.Replace(",", string.Empty, StringComparison.Ordinal).Replace(" ", string.Empty, StringComparison.Ordinal);
        return double.TryParse(numericText, NumberStyles.Any, CultureInfo.InvariantCulture, out var value)
            ? value
            : 0d;
    }

    private static IEnumerable<double> ResolveRange(SheetDocument document, int activeTabIndex, string token)
    {
        var parts = token.Split(':', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2)
        {
            yield break;
        }

        var left = NormalizeRangeEndpoint(parts[0], null);
        var right = NormalizeRangeEndpoint(parts[1], left.tabName);
        if (!TryParseCell(left.cellRef, out var startRow, out var startCol) || !TryParseCell(right.cellRef, out var endRow, out var endCol))
        {
            yield break;
        }

        var targetTabIndex = activeTabIndex;
        if (!string.IsNullOrWhiteSpace(left.tabName))
        {
            var lookup = document.Tabs.FindIndex(x => string.Equals(x.Name, left.tabName, StringComparison.OrdinalIgnoreCase));
            if (lookup >= 0)
            {
                targetTabIndex = lookup;
            }
        }

        var minRow = Math.Min(startRow, endRow);
        var maxRow = Math.Max(startRow, endRow);
        var minCol = Math.Min(startCol, endCol);
        var maxCol = Math.Max(startCol, endCol);
        for (var row = minRow; row <= maxRow; row++)
        {
            for (var col = minCol; col <= maxCol; col++)
            {
                yield return ResolveCellNumber(document, targetTabIndex, ToInternalCellKey(row, col));
            }
        }
    }

    private static (string? tabName, string cellRef) NormalizeRangeEndpoint(string token, string? fallbackTabName)
    {
        var bangIndex = token.LastIndexOf('!');
        if (bangIndex > 0)
        {
            return (token[..bangIndex].Trim(), token[(bangIndex + 1)..].Trim());
        }

        return (fallbackTabName, token.Trim());
    }

    private static bool TryParseCell(string token, out int row, out int column)
    {
        row = 0;
        column = 0;
        var match = CellRefRegex.Match(token.Trim());
        if (!match.Success)
        {
            return false;
        }

        var letters = match.Groups[1].Value.ToUpperInvariant();
        var digits = match.Groups[2].Value;
        foreach (var ch in letters)
        {
            column = (column * 26) + (ch - 'A' + 1);
        }

        return int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out row);
    }

    private static string ToInternalCellKey(int row, int column) => $"R{row}C{column}";

    [GeneratedRegex(@"^([A-Z]+)(\d+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase)]
    private static partial Regex CellRefRegexFactory();

    [GeneratedRegex(@"^\s*((?:[A-Za-z0-9_ ]+!)?[A-Z]+\d+)\s*:\s*((?:[A-Za-z0-9_ ]+!)?[A-Z]+\d+)", RegexOptions.Compiled | RegexOptions.IgnoreCase)]
    private static partial Regex RangeRegexFactory();

    private static readonly Regex CellRefRegex = CellRefRegexFactory();
    private static readonly Regex RangeRegex = RangeRegexFactory();

    private sealed class ExpressionParser
    {
        private readonly string text;
        private readonly SheetDocument document;
        private readonly int activeTabIndex;
        private int index;

        public ExpressionParser(string text, SheetDocument document, int activeTabIndex)
        {
            this.text = text;
            this.document = document;
            this.activeTabIndex = activeTabIndex;
        }

        public double ParseExpressionValue()
        {
            var value = this.ParseExpression();
            this.SkipWhitespace();
            if (this.index < this.text.Length)
            {
                throw new InvalidOperationException("Unexpected trailing characters.");
            }

            return value;
        }

        private double ParseExpression()
        {
            var left = this.ParseTerm();
            while (true)
            {
                this.SkipWhitespace();
                if (this.Match('+'))
                {
                    left += this.ParseTerm();
                    continue;
                }

                if (this.Match('-'))
                {
                    left -= this.ParseTerm();
                    continue;
                }

                return left;
            }
        }

        private double ParseTerm()
        {
            var left = this.ParseFactor();
            while (true)
            {
                this.SkipWhitespace();
                if (this.Match('*'))
                {
                    left *= this.ParseFactor();
                    continue;
                }

                if (this.Match('/'))
                {
                    var right = this.ParseFactor();
                    left = Math.Abs(right) < 0.0000001d ? 0d : left / right;
                    continue;
                }

                return left;
            }
        }

        private double ParseFactor()
        {
            this.SkipWhitespace();
            if (this.Match('+'))
            {
                return this.ParseFactor();
            }

            if (this.Match('-'))
            {
                return -this.ParseFactor();
            }

            if (this.Match('('))
            {
                var inner = this.ParseExpression();
                this.Expect(')');
                return inner;
            }

            if (this.TryParseFunction(out var functionValue))
            {
                return functionValue;
            }

            if (this.TryParseReference(out var referenceValue))
            {
                return referenceValue;
            }

            return this.ParseNumber();
        }

        private bool TryParseFunction(out double value)
        {
            value = 0d;
            var start = this.index;
            if (start >= this.text.Length || !char.IsLetter(this.text[start]))
            {
                return false;
            }

            while (this.index < this.text.Length && (char.IsLetter(this.text[this.index]) || this.text[this.index] == '_'))
            {
                this.index++;
            }

            var name = this.text[start..this.index].Trim();
            this.SkipWhitespace();
            if (!this.Match('('))
            {
                this.index = start;
                return false;
            }

            var args = new List<double>();
            this.SkipWhitespace();
            if (!this.Match(')'))
            {
                while (true)
                {
                    this.SkipWhitespace();
                    if (this.TryConsumeRangeValues(out var rangeValues))
                    {
                        args.AddRange(rangeValues);
                    }
                    else
                    {
                        args.Add(this.ParseExpression());
                    }

                    this.SkipWhitespace();
                    if (this.Match(')'))
                    {
                        break;
                    }

                    this.Expect(',');
                }
            }

            value = name.ToUpperInvariant() switch
            {
                "SUM" => args.Sum(),
                "AVG" or "AVERAGE" => args.Count == 0 ? 0d : args.Average(),
                "MIN" => args.Count == 0 ? 0d : args.Min(),
                "MAX" => args.Count == 0 ? 0d : args.Max(),
                "COUNT" => args.Count,
                _ => throw new InvalidOperationException("Unknown function."),
            };
            return true;
        }

        private bool TryConsumeRangeValues(out List<double> values)
        {
            values = new List<double>();
            var match = RangeRegex.Match(this.text[this.index..]);
            if (!match.Success || match.Index != 0)
            {
                return false;
            }

            var token = match.Value;
            values = ResolveRange(this.document, this.activeTabIndex, token).ToList();
            this.index += token.Length;
            return true;
        }

        private bool TryParseReference(out double value)
        {
            value = 0d;
            var match = Regex.Match(this.text[this.index..], @"^((?:[A-Za-z0-9_ ]+!)?[A-Z]+\d+)", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return false;
            }

            var token = match.Groups[1].Value;
            value = ResolveCellNumber(this.document, this.activeTabIndex, token);
            this.index += token.Length;
            return true;
        }

        private double ParseNumber()
        {
            this.SkipWhitespace();
            var start = this.index;
            var seenDecimal = false;
            while (this.index < this.text.Length)
            {
                var ch = this.text[this.index];
                if (char.IsDigit(ch))
                {
                    this.index++;
                    continue;
                }

                if ((ch == '.' || ch == ',') && !seenDecimal)
                {
                    seenDecimal = true;
                    this.index++;
                    continue;
                }

                break;
            }

            if (start == this.index)
            {
                throw new InvalidOperationException("Expected number.");
            }

            var slice = this.text[start..this.index].Replace(',', '.');
            return double.TryParse(slice, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
                ? value
                : 0d;
        }

        private void SkipWhitespace()
        {
            while (this.index < this.text.Length && char.IsWhiteSpace(this.text[this.index]))
            {
                this.index++;
            }
        }

        private bool Match(char ch)
        {
            if (this.index < this.text.Length && this.text[this.index] == ch)
            {
                this.index++;
                return true;
            }

            return false;
        }

        private void Expect(char ch)
        {
            this.SkipWhitespace();
            if (!this.Match(ch))
            {
                throw new InvalidOperationException($"Expected '{ch}'.");
            }
        }
    }
}
