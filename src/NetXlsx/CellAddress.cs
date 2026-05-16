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

    private static bool IsAsciiLetter(char c) =>
        c is (>= 'A' and <= 'Z') or (>= 'a' and <= 'z');
}
