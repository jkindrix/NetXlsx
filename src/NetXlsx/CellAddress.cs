// A1 cell-address parser per design §6.10.
// Single-cell form only in v0.2.0; range parsing lands with IRange in a
// follow-up milestone. Accepts: A1, a1 (case-insensitive), $A$1 / $A1 /
// A$1 ($ stripped). Rejects: Sheet1!A1, A:A, 1:1, empty, anything else.

using System;

namespace NetXlsx;

/// <summary>
/// Parses and formats Excel <c>A1</c>-style cell addresses. All integer
/// row and column indices are <strong>1-based</strong> (decision #3).
/// </summary>
public static class CellAddress
{
    /// <summary>Maximum row index (Excel hard limit, decision #44).</summary>
    public const int MaxRow = 1_048_576;
    /// <summary>Maximum column index — column "XFD" (decision #44).</summary>
    public const int MaxColumn = 16_384;

    /// <summary>
    /// Parses an <c>A1</c>-style cell reference. Returns 1-based row and
    /// column. Throws <see cref="InvalidCellAddressException"/> on failure.
    /// </summary>
    /// <param name="a1">A cell reference like <c>"A1"</c>, <c>"$A$1"</c>, or <c>"aa10"</c>.</param>
    public static (int Row, int Column) Parse(string a1)
    {
        if (!TryParse(a1, out int row, out int col, out string? reason))
        {
            throw new InvalidCellAddressException(a1 ?? "<null>", reason);
        }
        return (row, col);
    }

    /// <summary>
    /// Non-throwing parse. See <see cref="Parse"/>.
    /// </summary>
    public static bool TryParse(string a1, out int row, out int column)
    {
        return TryParse(a1, out row, out column, out _);
    }

    private static bool TryParse(string? a1, out int row, out int column, out string reason)
    {
        row = 0;
        column = 0;
        if (a1 is null)
        {
            reason = "input is null";
            return false;
        }
        if (a1.Length == 0)
        {
            reason = "input is empty";
            return false;
        }
        if (a1.Contains('!', StringComparison.Ordinal))
        {
            reason = "sheet-qualified references (e.g. \"Sheet1!A1\") are not accepted by the cell indexer; use the sheet's indexer directly";
            return false;
        }
        if (a1.Contains(':', StringComparison.Ordinal))
        {
            reason = "range references (e.g. \"A1:C10\", \"A:A\", \"1:1\") are not accepted by the cell indexer; use Range(...) instead";
            return false;
        }

        int i = 0;

        // Optional leading $ on column
        if (i < a1.Length && a1[i] == '$') i++;

        // Column letters
        int col = 0;
        int letterStart = i;
        while (i < a1.Length && IsAsciiLetter(a1[i]))
        {
            char c = a1[i];
            int v = (c >= 'a' ? c - 'a' : c - 'A') + 1;   // A=1, B=2, ...
            col = col * 26 + v;
            if (col > MaxColumn)
            {
                reason = $"column part exceeds Excel maximum (XFD = {MaxColumn})";
                return false;
            }
            i++;
        }
        if (i == letterStart)
        {
            reason = "missing column letters";
            return false;
        }

        // Optional $ between column and row
        if (i < a1.Length && a1[i] == '$') i++;

        // Row digits
        int rowStart = i;
        long rowAcc = 0;
        while (i < a1.Length && a1[i] is >= '0' and <= '9')
        {
            rowAcc = rowAcc * 10 + (a1[i] - '0');
            if (rowAcc > MaxRow)
            {
                reason = $"row part exceeds Excel maximum ({MaxRow})";
                return false;
            }
            i++;
        }
        if (i == rowStart)
        {
            reason = "missing row digits";
            return false;
        }
        if (rowAcc == 0)
        {
            reason = "row index must be 1 or greater";
            return false;
        }

        if (i != a1.Length)
        {
            reason = $"unexpected trailing characters '{a1.Substring(i)}'";
            return false;
        }

        row = (int)rowAcc;
        column = col;
        reason = "";
        return true;
    }

    /// <summary>
    /// Parses an Excel range reference. Returns 1-based top-left and
    /// bottom-right coordinates. Accepts:
    /// <list type="bullet">
    /// <item>Single cell — <c>A1</c>, <c>$A$1</c>, <c>aa10</c>.</item>
    /// <item>Bounded range — <c>A1:C3</c>, with case + <c>$</c> tolerance.</item>
    /// <item>Whole-column shorthand — <c>A:A</c> or <c>A:C</c>, auto-expanded
    /// to <c>A1:A1048576</c> / <c>A1:C1048576</c> respectively.</item>
    /// <item>Whole-row shorthand — <c>1:1</c> or <c>1:5</c>, auto-expanded
    /// to <c>A1:XFD1</c> / <c>A1:XFD5</c>.</item>
    /// </list>
    /// Both sides of the colon must use the same form
    /// (cell+cell / column+column / row+row); mixed forms throw.
    /// </summary>
    /// <exception cref="InvalidCellAddressException">The string is not a valid range.</exception>
    public static (int Row1, int Col1, int Row2, int Col2) ParseRange(string a1Range)
    {
        if (a1Range is null)
            throw new InvalidCellAddressException("<null>", "input is null");

        int colonIdx = a1Range.IndexOf(':');
        if (colonIdx < 0)
        {
            // Single cell — treat as a 1x1 range.
            var (r, c) = Parse(a1Range);
            return (r, c, r, c);
        }

        var left = a1Range.Substring(0, colonIdx);
        var right = a1Range.Substring(colonIdx + 1);

        // Both sides single-cell? (A1:C3 form.)
        if (TryParse(left, out int leftRow, out int leftCol, out _)
            && TryParse(right, out int rightRow, out int rightCol, out _))
        {
            if (leftRow > rightRow) (leftRow, rightRow) = (rightRow, leftRow);
            if (leftCol > rightCol) (leftCol, rightCol) = (rightCol, leftCol);
            return (leftRow, leftCol, rightRow, rightCol);
        }

        // Both sides column-only? (A:A or A:C — whole-column form.)
        if (TryParseColumnOnly(left, out int leftColOnly)
            && TryParseColumnOnly(right, out int rightColOnly))
        {
            if (leftColOnly > rightColOnly) (leftColOnly, rightColOnly) = (rightColOnly, leftColOnly);
            return (1, leftColOnly, MaxRow, rightColOnly);
        }

        // Both sides row-only? (1:1 or 1:5 — whole-row form.)
        if (TryParseRowOnly(left, out int leftRowOnly)
            && TryParseRowOnly(right, out int rightRowOnly))
        {
            if (leftRowOnly > rightRowOnly) (leftRowOnly, rightRowOnly) = (rightRowOnly, leftRowOnly);
            return (leftRowOnly, 1, rightRowOnly, MaxColumn);
        }

        throw new InvalidCellAddressException(a1Range,
            "range sides must both be single cells (A1:C3), both columns (A:C), or both rows (1:5)");
    }

    /// <summary>
    /// True when <paramref name="input"/> is a column-letter reference
    /// with no row digits — <c>A</c>, <c>AA</c>, <c>$A</c>. Out param is
    /// the 1-based column index.
    /// </summary>
    private static bool TryParseColumnOnly(string input, out int column)
    {
        column = 0;
        if (string.IsNullOrEmpty(input)) return false;
        int i = 0;
        if (input[i] == '$') i++;
        int col = 0;
        int letterStart = i;
        while (i < input.Length && IsAsciiLetter(input[i]))
        {
            char c = input[i];
            int v = (c >= 'a' ? c - 'a' : c - 'A') + 1;
            col = col * 26 + v;
            if (col > MaxColumn) return false;
            i++;
        }
        if (i == letterStart) return false;
        if (i != input.Length) return false;   // anything trailing → not column-only
        column = col;
        return true;
    }

    /// <summary>
    /// True when <paramref name="input"/> is a row-number reference with
    /// no column letters — <c>1</c>, <c>10</c>, <c>$5</c>. Out param is
    /// the 1-based row index.
    /// </summary>
    private static bool TryParseRowOnly(string input, out int row)
    {
        row = 0;
        if (string.IsNullOrEmpty(input)) return false;
        int i = 0;
        if (input[i] == '$') i++;
        long acc = 0;
        int digitStart = i;
        while (i < input.Length && input[i] is >= '0' and <= '9')
        {
            acc = acc * 10 + (input[i] - '0');
            if (acc > MaxRow) return false;
            i++;
        }
        if (i == digitStart) return false;
        if (i != input.Length) return false;
        if (acc == 0) return false;
        row = (int)acc;
        return true;
    }

    /// <summary>
    /// Formats a 1-based range as the canonical <c>A1:C3</c> string.
    /// Degenerate 1x1 ranges round-trip as the single-cell form
    /// (<c>A1</c>) so callers see the canonical address.
    /// </summary>
    public static string FormatRange(int row1, int col1, int row2, int col2)
    {
        if (row1 == row2 && col1 == col2) return Format(row1, col1);
        return $"{Format(row1, col1)}:{Format(row2, col2)}";
    }

    /// <summary>
    /// Formats a 1-based row and column as the canonical <c>A1</c> string
    /// — uppercase column letters, no <c>$</c>.
    /// </summary>
    public static string Format(int row, int column)
    {
        if (row < 1 || row > MaxRow)
            throw new ArgumentOutOfRangeException(nameof(row), row, $"row must be in [1, {MaxRow}]");
        if (column < 1 || column > MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(column), column, $"column must be in [1, {MaxColumn}]");

        // Excel's column lettering is bijective base-26: A..Z, AA..AZ, BA..ZZ, AAA..XFD.
        // Standard algorithm: subtract 1 each iteration.
        Span<char> buf = stackalloc char[8];
        int idx = buf.Length;
        int c = column;
        while (c > 0)
        {
            c--;
            buf[--idx] = (char)('A' + c % 26);
            c /= 26;
        }
        return string.Concat(buf.Slice(idx).ToString(), row.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Parses a column-letter reference (e.g. <c>"A"</c>, <c>"AA"</c>,
    /// <c>"XFD"</c>) into its 1-based column index. A leading <c>$</c>
    /// (absolute marker) is accepted and ignored. Case-insensitive.
    /// </summary>
    /// <exception cref="InvalidCellAddressException">Not a valid column letter.</exception>
    public static int ParseColumn(string letter)
    {
        ArgumentNullException.ThrowIfNull(letter);
        if (!TryParseColumnOnly(letter, out int col))
            throw new InvalidCellAddressException(letter,
                $"not a valid column letter (expected A..XFD)");
        return col;
    }

    /// <summary>
    /// Non-throwing form of <see cref="ParseColumn"/>.
    /// </summary>
    public static bool TryParseColumn(string letter, out int column) =>
        TryParseColumnOnly(letter, out column);

    /// <summary>
    /// Formats a 1-based column index as its letter form (<c>1 → "A"</c>,
    /// <c>27 → "AA"</c>).
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="column"/> outside <c>[1, MaxColumn]</c>.</exception>
    public static string FormatColumn(int column)
    {
        if (column < 1 || column > MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(column), column, $"column must be in [1, {MaxColumn}]");
        Span<char> buf = stackalloc char[8];
        int idx = buf.Length;
        int c = column;
        while (c > 0)
        {
            c--;
            buf[--idx] = (char)('A' + c % 26);
            c /= 26;
        }
        return buf.Slice(idx).ToString();
    }

    private static bool IsAsciiLetter(char c) =>
        c is (>= 'A' and <= 'Z') or (>= 'a' and <= 'z');
}
