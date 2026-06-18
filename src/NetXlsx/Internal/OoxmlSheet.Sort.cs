// I-82 engine swap — Open XML SDK-backed SortRange (CF/validation/tables/
// autofilter/sort slice).
//
// SortRange (decision I-72) is an in-memory grid operation: rows within the
// range are physically reordered; no OOXML <sortState> metadata is written
// (matching the NPOI engine). The comparison semantics mirror the NPOI
// engine's CellSnapshot.Compare exactly — blanks sort last, numbers before
// strings (Excel's default), strings compare ordinal-ignore-case, FALSE <
// TRUE — and the sort is stable: an index permutation is sorted with the
// original row index as the final tie-break (lesson #12: Excel's sort is
// stable).
//
// Where the NPOI engine re-applies captured values to cells, this engine
// detaches the in-range <c> elements and re-homes them into their target
// rows (updating @r). Moving the element rather than a value snapshot
// carries the style index, inline rich text, and formula text verbatim by
// construction — including the documented formula caveat (lesson #12 /
// ISheet.SortRange: formula cells keep their literal text; relative
// references are NOT relocated).

using System;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    public void SortRange(string a1Range, params SortKey[] keys)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(a1Range);
        ArgumentNullException.ThrowIfNull(keys);
        if (keys.Length == 0) throw new ArgumentException("At least one sort key is required.", nameof(keys));

        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);

        int rowCount = r2 - r1 + 1;
        if (rowCount <= 1) return;

        int colCount = c2 - c1 + 1;

        // Snapshot comparison keys (from the live DOM, before any mutation)
        // and the backing <c> elements per row.
        var compare = new SortCellKey[rowCount][];
        var elements = new S.Cell?[rowCount][];
        for (int ri = 0; ri < rowCount; ri++)
        {
            compare[ri] = new SortCellKey[colCount];
            elements[ri] = new S.Cell?[colCount];
            for (int ci = 0; ci < colCount; ci++)
            {
                elements[ri][ci] = FindCell(r1 + ri, c1 + ci);
                compare[ri][ci] = SortCellKey.Capture(this, r1 + ri, c1 + ci);
            }
        }

        // Sort rows by keys. Excel's sort is stable, so we sort an index
        // permutation and break full-key ties by the original row index —
        // making the otherwise-unstable Array.Sort stable. Rows that tie
        // on every key keep their relative order.
        var order = new int[rowCount];
        for (int i = 0; i < rowCount; i++) order[i] = i;
        Array.Sort(order, (ia, ib) =>
        {
            var a = compare[ia];
            var b = compare[ib];
            foreach (var key in keys)
            {
                int colIdx = key.Column - c1;
                if (colIdx < 0 || colIdx >= colCount) continue;
                int cmp = SortCellKey.Compare(a[colIdx], b[colIdx]);
                if (cmp != 0) return key.Ascending ? cmp : -cmp;
            }
            return ia.CompareTo(ib);
        });

        // Detach every in-range cell element, then re-home each into its
        // target row in permutation order. <c> elements within a row must
        // stay in ascending column order; cells outside the sorted column
        // span may interleave, so each re-insert positions before the first
        // higher-column sibling.
        foreach (var row in elements)
            foreach (var c in row)
                c?.Remove();

        for (int ri = 0; ri < rowCount; ri++)
        {
            var source = elements[order[ri]];
            for (int ci = 0; ci < colCount; ci++)
            {
                var c = source[ci];
                if (c is null) continue;
                c.CellReference = CellAddress.Format(r1 + ri, c1 + ci);
                InsertCellInColumnOrder(GetOrCreateRow(r1 + ri), c, c1 + ci);
            }
        }
    }

    // Inserts a detached <c> element into a row, keeping <c> children in
    // ascending column order (the same invariant GetOrCreateCell maintains
    // for newly created cells).
    private static void InsertCellInColumnOrder(S.Row row, S.Cell cell, int col)
    {
        S.Cell? successor = null;
        foreach (var existing in row.Elements<S.Cell>())
        {
            if (ColumnOf(existing) > col) { successor = existing; break; }
        }
        if (successor is null) row.AppendChild(cell);
        else row.InsertBefore(cell, successor);
    }

    // A cell's sort-comparison key. Mirrors the NPOI engine's CellSnapshot
    // compare axes: a kind class (Date folds into Number — a date is a
    // numeric serial), plus the value for the comparable kinds. Formula and
    // Error cells always tie (the NPOI engine's comparator reaches its
    // "return 0" fall-through for them), so the stable tie-break preserves
    // their relative order.
    private readonly struct SortCellKey
    {
        public CellKind Kind { get; private init; }
        public double Num { get; private init; }
        public string? Str { get; private init; }
        public bool Bool { get; private init; }

        public static SortCellKey Capture(OoxmlSheet sheet, int row, int col)
        {
            var cell = sheet.CellHandle(row, col);
            var kind = cell.Kind;
            if (kind == CellKind.Date) kind = CellKind.Number;
            return kind switch
            {
                CellKind.Number => new SortCellKey { Kind = kind, Num = cell.GetNumber()!.Value },
                CellKind.String => new SortCellKey { Kind = kind, Str = cell.GetString() },
                CellKind.Bool => new SortCellKey { Kind = kind, Bool = cell.GetBool()!.Value },
                _ => new SortCellKey { Kind = kind },
            };
        }

        public static int Compare(in SortCellKey a, in SortCellKey b)
        {
            // Blanks sort last
            if (a.Kind == CellKind.Empty && b.Kind == CellKind.Empty) return 0;
            if (a.Kind == CellKind.Empty) return 1;
            if (b.Kind == CellKind.Empty) return -1;

            if (a.Kind == CellKind.Number && b.Kind == CellKind.Number)
                return a.Num.CompareTo(b.Num);

            if (a.Kind == CellKind.String && b.Kind == CellKind.String)
                return StringComparer.OrdinalIgnoreCase.Compare(a.Str, b.Str);

            // Mixed: numbers sort before strings (Excel default)
            if (a.Kind == CellKind.Number) return -1;
            if (b.Kind == CellKind.Number) return 1;

            // Booleans: FALSE < TRUE
            if (a.Kind == CellKind.Bool && b.Kind == CellKind.Bool)
                return a.Bool.CompareTo(b.Bool);

            return 0;
        }
    }
}
