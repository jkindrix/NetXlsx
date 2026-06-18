// I-82 engine swap — Open XML SDK-backed merged regions (structure slice).
//
// OOXML stores merged regions as <mergeCells count="N"><mergeCell ref="A1:C3"/>…
// </mergeCells>, a worksheet child that sits AFTER <sheetData> in CT_Worksheet's
// strict child sequence (SDK-quirk #3 — AppendChild does not reorder). We insert
// it immediately after <sheetData>, mirroring how OoxmlSheet inserts <cols>
// before <sheetData>. The schema-validation gate (slice 4b) holds this ordering
// to OpenXmlValidator on every merge fixture.
//
// CT_MergeCells requires at least one <mergeCell> child, so an emptied
// <mergeCells> is removed entirely rather than left childless. The optional
// @count attribute is kept in sync with the child count (Excel and NPOI both
// emit it).
//
// Contract mirrors the NPOI engine (XssfSheet): a 1x1 merge is a no-op
// (decision I-38); an overlapping merge throws InvalidOperationException
// (design §6.4); UnmergeCells of a non-existent exact range is a silent no-op;
// MergedRanges returns canonical A1:C3 strings. Lesson #4 (merged-region borders
// render from the boundary cells) is honored by MergeCellsStyled, which styles
// every cell in the range before merging — and emits cells before the merge.

using System;
using System.Collections.Generic;
using System.Linq;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    public void MergeCells(string a1Range)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(a1Range);
        // ParseRange already normalizes corners (r1<=r2, c1<=c2).
        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);

        // 1x1 merge is a no-op per decision I-38.
        if (r1 == r2 && c1 == c2) return;

        var existing = Worksheet.GetFirstChild<S.MergeCells>();
        if (existing is not null)
        {
            foreach (var mc in existing.Elements<S.MergeCell>())
            {
                var (er1, ec1, er2, ec2) = CellAddress.ParseRange(mc.Reference!.Value!);
                if (RangesOverlap(r1, c1, r2, c2, er1, ec1, er2, ec2))
                    throw new InvalidOperationException(
                        $"MergeCells('{a1Range}') overlaps existing merged region " +
                        $"{CellAddress.FormatRange(er1, ec1, er2, ec2)}.");
            }
        }

        var container = GetOrCreateMergeCells();
        container.AppendChild(new S.MergeCell { Reference = CellAddress.FormatRange(r1, c1, r2, c2) });
        container.Count = (uint)container.Elements<S.MergeCell>().Count();
    }

    public void MergeCellsStyled(string a1Range, CellStyle style)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(a1Range);
        ArgumentNullException.ThrowIfNull(style);
        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);
        // Style every cell in the range so borders render across the full merged
        // area (lesson #4 / decision I-79). Cells are emitted here, before the
        // merge is added below — the OOXML "cells before merges" ordering.
        for (int r = r1; r <= r2; r++)
            for (int c = c1; c <= c2; c++)
                this[r, c].Style(style);
        if (r1 != r2 || c1 != c2) MergeCells(a1Range);
    }

    public void UnmergeCells(string a1Range)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(a1Range);
        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);

        var merges = Worksheet.GetFirstChild<S.MergeCells>();
        if (merges is null) return;

        foreach (var mc in merges.Elements<S.MergeCell>().ToList())
        {
            var (er1, ec1, er2, ec2) = CellAddress.ParseRange(mc.Reference!.Value!);
            if (er1 == r1 && ec1 == c1 && er2 == r2 && ec2 == c2)
            {
                mc.Remove();
                break; // exact-match region is unique
            }
        }

        int remaining = merges.Elements<S.MergeCell>().Count();
        // CT_MergeCells requires >= 1 child — drop the container when empty
        // rather than leave a schema-invalid childless <mergeCells/>.
        if (remaining == 0) merges.Remove();
        else merges.Count = (uint)remaining;
    }

    public IReadOnlyList<string> MergedRanges
    {
        get
        {
            ThrowIfUnusable();
            var merges = Worksheet.GetFirstChild<S.MergeCells>();
            if (merges is null) return Array.Empty<string>();
            var list = new List<string>();
            foreach (var mc in merges.Elements<S.MergeCell>())
            {
                // Re-parse + re-format so a file-authored ref is returned canonically.
                var (r1, c1, r2, c2) = CellAddress.ParseRange(mc.Reference!.Value!);
                list.Add(CellAddress.FormatRange(r1, c1, r2, c2));
            }
            return list;
        }
    }

    // <mergeCells> sits after <sheetData> (and after <sheetProtection> /
    // <autoFilter> / <sortState> / … when present) in CT_Worksheet's child
    // sequence. OoxmlSchemaOrder places it correctly even on an opened file that
    // already carries an intervening sibling (SDK-quirk #8); a bare
    // InsertAfter(<sheetData>) would emit out-of-order XML in that case.
    private S.MergeCells GetOrCreateMergeCells()
        => OoxmlSchemaOrder.GetOrInsert(Worksheet, static () => new S.MergeCells());

    private static bool RangesOverlap(
        int ar1, int ac1, int ar2, int ac2,
        int br1, int bc1, int br2, int bc2) =>
        ar1 <= br2 && br1 <= ar2 && ac1 <= bc2 && bc1 <= ac2;
}
