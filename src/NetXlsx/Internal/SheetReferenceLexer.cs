// I-90 sheet lifecycle — the sheet-reference lexer (S2 memo, signed off as
// amended 2026-06-11).
//
// Rewrites references to a renamed sheet in formula-shaped text WITHOUT a
// formula parser: recognizing the 'Quoted Name'! and BareName! prefixes only
// requires tracking "inside a string literal" / "inside a quoted sheet name"
// — the same two literal forms SetFormula's structural validator already
// lexes (OoxmlCell.ValidateFormulaStructure). Callers run it over cell
// formulas, defined-name bodies, CF/DV formulas, chart c:f, sparkline xm:f,
// table column formulas, and internal hyperlink @location values.
//
// Matching is case-insensitive (sheet-name semantics, the engine-wide
// OrdinalIgnoreCase rule). Output quoting is normalized per the memo: the
// new name is quoted iff it needs quoting, with embedded apostrophes
// doubled. Format quotes strictly more often than the chart/autofilter
// writers' QuoteSheetName (cell-reference-shaped and boolean-literal names
// are quoted here) — over-quoting is always a valid formula; those writers'
// output is pinned against the NPOI oracle and stays untouched.
//
// Documented residuals:
//   * Sheet names inside string arguments (INDIRECT("Old!A1")) are never
//     rewritten — Excel parity; Excel does not rewrite them either.
//   * 3D span references (Sheet1:Sheet3!A1 and 'First:Last'!A1) are NOT
//     rewritten. This diverges from Excel, which follows both endpoints.
//     The ':' predecessor guard keeps a bare 3D endpoint from being
//     rewritten into the malformed First:'New Name'! shape; the quoted 3D
//     pair can never match because a legal sheet name cannot contain ':'.
//   * External-workbook references ([1]Sheet1!A1) are skipped via the ']'
//     predecessor guard; the quoted form ('[Book1]Sheet1'!) never matches
//     because '[' is illegal in a sheet name.
//   * Error literals (#REF!A1) are protected by the '#' predecessor guard.

using System;
using System.Text;

namespace NetXlsx;

internal static class SheetReferenceLexer
{
    /// <summary>
    /// Rewrites every <c>'oldName'!</c> / <c>oldName!</c> sheet-reference
    /// prefix outside string literals to the (normalized-quoted) new name.
    /// Returns the original string instance when nothing matched, so callers
    /// can skip the DOM write via reference equality.
    /// </summary>
    internal static string Rewrite(string text, string oldName, string newName)
    {
        StringBuilder? sb = null;
        int copied = 0;
        int i = 0;
        int n = text.Length;
        while (i < n)
        {
            char ch = text[i];
            if (ch == '"')
            {
                i = SkipStringLiteral(text, i);
                continue;
            }
            if (ch == '\'')
            {
                int contentStart = i + 1;
                if (!TryScanQuoted(text, contentStart, out int contentEnd, out int after))
                    break; // unterminated quoted literal — leave the tail untouched
                if (after < n && text[after] == '!'
                    && QuotedContentEquals(text, contentStart, contentEnd, oldName))
                {
                    sb ??= new StringBuilder(n + 16);
                    sb.Append(text, copied, i - copied);
                    sb.Append(Format(newName));
                    copied = after; // keep the '!' and everything after
                }
                i = after;
                continue;
            }
            if (IsBareNameChar(ch))
            {
                int start = i;
                do { i++; } while (i < n && IsBareNameChar(text[i]));
                if (i < n && text[i] == '!'
                    && (start == 0 || !IsBareRunGuard(text[start - 1]))
                    && i - start == oldName.Length
                    && string.Compare(text, start, oldName, 0, oldName.Length, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    sb ??= new StringBuilder(n + 16);
                    sb.Append(text, copied, start - copied);
                    sb.Append(Format(newName));
                    copied = i;
                }
                continue;
            }
            i++;
        }
        if (sb is null) return text;
        sb.Append(text, copied, n - copied);
        return sb.ToString();
    }

    /// <summary>
    /// Formats a sheet name for embedding in a formula: quoted (with
    /// embedded apostrophes doubled) iff the bare form would be ambiguous.
    /// </summary>
    internal static string Format(string name)
        => NeedsQuoting(name) ? "'" + name.Replace("'", "''") + "'" : name;

    // A name can appear bare only when it is an identifier-shaped token that
    // cannot be confused with anything else the formula grammar puts before
    // '!': letters/digits/underscore, not digit-leading, not an A1- or
    // R1C1-style cell reference, not a boolean literal.
    private static bool NeedsQuoting(string name)
    {
        if (name.Length == 0 || char.IsDigit(name[0])) return true;
        foreach (char c in name)
            if (!(char.IsLetterOrDigit(c) || c == '_')) return true;
        return LooksLikeA1Reference(name)
            || LooksLikeR1C1Reference(name)
            || name.Equals("TRUE", StringComparison.OrdinalIgnoreCase)
            || name.Equals("FALSE", StringComparison.OrdinalIgnoreCase);
    }

    // 1-3 ASCII letters followed by only digits (A1, AB12, XFD1048576) —
    // the same conservative shape OoxmlWorkbook.LooksLikeCellReference pins
    // for defined names. Grid bounds are deliberately not checked:
    // over-quoting is safe.
    private static bool LooksLikeA1Reference(string name)
    {
        int i = 0;
        while (i < name.Length && IsAsciiLetter(name[i])) i++;
        if (i == 0 || i > 3 || i == name.Length) return false;
        for (int j = i; j < name.Length; j++)
            if (name[j] is < '0' or > '9') return false;
        return true;
    }

    // R, C, R1, C1, RC, R1C1, R1C, RC1 — names Excel quotes because the
    // bare form collides with R1C1-style references.
    private static bool LooksLikeR1C1Reference(string name)
    {
        int i = 0;
        bool any = false;
        if (i < name.Length && name[i] is 'R' or 'r')
        {
            i++; any = true;
            while (i < name.Length && name[i] is >= '0' and <= '9') i++;
        }
        if (i < name.Length && name[i] is 'C' or 'c')
        {
            i++; any = true;
            while (i < name.Length && name[i] is >= '0' and <= '9') i++;
        }
        return any && i == name.Length;
    }

    private static bool IsAsciiLetter(char c) => c is (>= 'A' and <= 'Z') or (>= 'a' and <= 'z');

    // The token characters a bare sheet reference can be built from. '.' is
    // accepted on the MATCH side (a producer may emit Ver1.2!A1 bare, and
    // including it keeps a run like Sheet1.x from false-matching its
    // Sheet1 prefix) even though Format never EMITS a bare period.
    private static bool IsBareNameChar(char c)
        => char.IsLetterOrDigit(c) || c is '_' or '.';

    // Characters that disqualify the run before '!' from being a plain
    // same-workbook sheet reference: '#' (error literals, #REF!A1), ']'
    // (external-workbook refs, [1]Sheet1!A1), ':' (a 3D span endpoint —
    // documented residual), and '!' / quote adjacency (malformed input;
    // never produced by a valid formula).
    private static bool IsBareRunGuard(char prev)
        => prev is '#' or ']' or ':' or '!' or '\'' or '"';

    // i sits on the opening '"'; returns the index just past the closing
    // quote ("" is the escape form). An unterminated literal swallows the
    // tail — matching the structural validator's model of the same text.
    private static int SkipStringLiteral(string text, int i)
    {
        i++;
        while (i < text.Length)
        {
            if (text[i] == '"')
            {
                if (i + 1 < text.Length && text[i + 1] == '"') { i += 2; continue; }
                return i + 1;
            }
            i++;
        }
        return i;
    }

    // Scans a quoted run from the char after the opening apostrophe.
    // contentEnd is the index of the closing apostrophe (content excludes
    // it); after is the index just past it. Doubled apostrophes are content.
    private static bool TryScanQuoted(string text, int contentStart, out int contentEnd, out int after)
    {
        int i = contentStart;
        while (i < text.Length)
        {
            if (text[i] == '\'')
            {
                if (i + 1 < text.Length && text[i + 1] == '\'') { i += 2; continue; }
                contentEnd = i;
                after = i + 1;
                return true;
            }
            i++;
        }
        contentEnd = after = text.Length;
        return false;
    }

    // Compares the quoted run's content (with '' unescaped on the fly)
    // against a sheet name, ordinal-ignore-case.
    private static bool QuotedContentEquals(string text, int start, int end, string name)
    {
        int ti = start, ni = 0;
        while (ti < end)
        {
            char c = text[ti];
            // TryScanQuoted guarantees apostrophes inside the run are paired.
            ti += c == '\'' ? 2 : 1;
            if (ni >= name.Length) return false;
            if (char.ToUpperInvariant(c) != char.ToUpperInvariant(name[ni])) return false;
            ni++;
        }
        return ni == name.Length;
    }
}
