// I-82 engine swap — Open XML SDK-backed AutoFilter (CF/validation/tables/
// autofilter/sort slice). Surface per decisions I-56 + I-66.
//
// OOXML stores the filter as a single <autoFilter ref="…"> worksheet child
// (0..1) holding one <filterColumn colId="N"> per filtered column, each with a
// <customFilters [and="1"]> of 1-2 <customFilter operator="…" val="…"/>
// conditions. <autoFilter> sits between <sheetData> and <mergeCells> in
// CT_Worksheet's strict child sequence, so the insert routes through
// OoxmlSchemaOrder (SDK-quirk #8). It is a 0..1 singleton, so GetOrInsert's
// get-or-insert shape applies directly.
//
// Contract mirrors the NPOI engine (XssfSheet.AutoFilter.cs), oracle-checked
// against its emitted XML (SDK-quirk #11 habit):
//   - SetAutoFilter creates/updates Excel's hidden built-in
//     _xlnm._FilterDatabase defined name ('Sheet'!$A$1:$B$2, no 1x1 collapse)
//     scoped to this sheet — NPOI's XSSFSheet.SetAutoFilter does the same.
//   - Re-setting the range keeps existing <filterColumn> entries (NPOI
//     replaces only @ref) and updates the defined name.
//   - ClearAutoFilter removes the element but leaves the stale
//     _FilterDatabase name in place — Excel tolerates it (NPOI parity).
//   - NPOI also emits showButton="1" on each filterColumn; that is the OOXML
//     default value (a cosmetic), so this engine omits it.

using System;
using System.Linq;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    public void SetAutoFilter(string a1Range)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(a1Range);
        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);

        var af = OoxmlSchemaOrder.GetOrInsert(Worksheet, static () => new S.AutoFilter());
        // 1x1 collapses to "A1" (NPOI's CellRangeAddress.FormatAsString parity).
        af.Reference = CellAddress.FormatRange(r1, c1, r2, c2);

        // Excel's hidden built-in name always carries the r1:r2 form (no 1x1
        // collapse) with absolute markers and a quoted-when-needed sheet name.
        _workbook.SetFilterDatabaseName(
            _name, $"{QuoteSheetName(_name)}!{AbsoluteRef(r1, c1)}:{AbsoluteRef(r2, c2)}");
    }

    public void ClearAutoFilter()
    {
        ThrowIfUnusable();
        // The auxiliary _FilterDatabase built-in name is left in place — Excel
        // tolerates a stale name pointing at an absent autoFilter (NPOI parity).
        Worksheet.GetFirstChild<S.AutoFilter>()?.Remove();
    }

    public bool HasAutoFilter
    {
        get
        {
            ThrowIfUnusable();
            return Worksheet.GetFirstChild<S.AutoFilter>() is not null;
        }
    }

    public string? AutoFilterRange
    {
        get
        {
            ThrowIfUnusable();
            return Worksheet.GetFirstChild<S.AutoFilter>()?.Reference?.Value;
        }
    }

    // ---- Per-column criteria (decision I-66) --------------------------

    public void SetAutoFilterColumn(int columnOffset, FilterCriteria criteria)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(criteria);

        var af = Worksheet.GetFirstChild<S.AutoFilter>()
            ?? throw new InvalidOperationException(
                "No AutoFilter is set on this sheet. Call SetAutoFilter(range) first.");

        EnsureColumnInRange(af, columnOffset);

        // Replace any existing entry for this column.
        RemoveFilterColumns(af, columnOffset);

        var fc = new S.FilterColumn { ColumnId = (uint)columnOffset };
        fc.AppendChild(BuildCustomFilters(criteria));

        // CT_AutoFilter is a sequence: filterColumn* then sortState / extLst.
        // An opened file may carry those trailing siblings — insert before the
        // first non-filterColumn child rather than blindly appending.
        var successor = af.ChildElements.FirstOrDefault(e => e is not S.FilterColumn);
        if (successor is null) af.AppendChild(fc);
        else af.InsertBefore(fc, successor);
    }

    public void ClearAutoFilterColumn(int columnOffset)
    {
        ThrowIfUnusable();
        var af = Worksheet.GetFirstChild<S.AutoFilter>();
        if (af is null) return;
        RemoveFilterColumns(af, columnOffset);
    }

    private static void RemoveFilterColumns(S.AutoFilter af, int columnOffset)
    {
        foreach (var fc in af.Elements<S.FilterColumn>()
                     .Where(f => f.ColumnId?.Value == (uint)columnOffset).ToList())
            fc.Remove();
    }

    private static void EnsureColumnInRange(S.AutoFilter af, int columnOffset)
    {
        if (columnOffset < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(columnOffset), columnOffset, "columnOffset must be >= 0.");
        }
        // @ref is required by CT_AutoFilter; a missing/corrupt one surfaces as
        // the same parse failure the NPOI engine raises for a bad af.@ref.
        var (_, c1, _, c2) = CellAddress.ParseRange(af.Reference?.Value!);
        int rangeWidth = c2 - c1 + 1;
        if (columnOffset >= rangeWidth)
        {
            throw new ArgumentOutOfRangeException(
                nameof(columnOffset), columnOffset,
                $"columnOffset must be < AutoFilter range width ({rangeWidth} for '{af.Reference!.Value}').");
        }
    }

    private static S.CustomFilters BuildCustomFilters(FilterCriteria criteria)
    {
        var filters = new S.CustomFilters();
        // @and's OOXML default is false; NPOI leaves it unset for Single/Or too.
        if (criteria.Combine == FilterCriteria.Combinator.And) filters.And = true;
        filters.AppendChild(new S.CustomFilter { Operator = ToOperator(criteria.Op1), Val = criteria.Val1 });
        if (criteria.Combine != FilterCriteria.Combinator.Single)
            filters.AppendChild(new S.CustomFilter { Operator = ToOperator(criteria.Op2!.Value), Val = criteria.Val2! });
        return filters;
    }

    private static S.FilterOperatorValues ToOperator(FilterCriteria.Op op) => op switch
    {
        FilterCriteria.Op.Equal => S.FilterOperatorValues.Equal,
        FilterCriteria.Op.NotEqual => S.FilterOperatorValues.NotEqual,
        FilterCriteria.Op.GreaterThan => S.FilterOperatorValues.GreaterThan,
        FilterCriteria.Op.GreaterThanOrEqual => S.FilterOperatorValues.GreaterThanOrEqual,
        FilterCriteria.Op.LessThan => S.FilterOperatorValues.LessThan,
        FilterCriteria.Op.LessThanOrEqual => S.FilterOperatorValues.LessThanOrEqual,
        _ => S.FilterOperatorValues.Equal,
    };

    // ---- formula-text helpers ------------------------------------------

    // $A$1-style absolute reference for the _FilterDatabase formula body.
    private static string AbsoluteRef(int row, int col)
    {
        var plain = CellAddress.Format(row, col);
        int split = 0;
        while (split < plain.Length && !char.IsDigit(plain[split])) split++;
        return $"${plain[..split]}${plain[split..]}";
    }

    // Quotes a sheet name for use in a formula body when it isn't a simple
    // identifier (letters/digits/underscore, not digit-leading) — e.g.
    // 'My Sheet'!$A$1:$B$2, with embedded apostrophes doubled (NPOI's
    // SheetNameFormatter behavior).
    private static string QuoteSheetName(string name)
    {
        bool simple = name.Length > 0 && !char.IsDigit(name[0]);
        if (simple)
        {
            foreach (var ch in name)
                if (!(char.IsLetterOrDigit(ch) || ch == '_')) { simple = false; break; }
        }
        return simple ? name : $"'{name.Replace("'", "''")}'";
    }
}
