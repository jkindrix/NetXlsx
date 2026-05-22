// XssfSheet — AutoFilter surface per decisions I-56 + I-66.
// Core class structure is in XssfSheet.cs.

using System;
using System.Collections.Generic;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Util;

namespace NetXlsx;

internal sealed partial class XssfSheet
{
    public void SetAutoFilter(string a1Range)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(a1Range);
        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);
        _underlying.SetAutoFilter(new CellRangeAddress(r1 - 1, r2 - 1, c1 - 1, c2 - 1));
    }

    public void ClearAutoFilter()
    {
        _workbook.ThrowIfDisposed();
        // NPOI 2.7.3 does not expose IsSet / unset accessors on
        // CT_Worksheet for the autoFilter element; setting the
        // property to null removes it from the serialized XML.
        // The auxiliary _FilterDatabase built-in name is left in
        // place — Excel tolerates a stale name pointing at an
        // absent autoFilter.
        _underlying.GetCTWorksheet().autoFilter = null;
    }

    public bool HasAutoFilter
    {
        get
        {
            _workbook.ThrowIfDisposed();
            return _underlying.GetCTWorksheet().autoFilter is not null;
        }
    }

    public string? AutoFilterRange
    {
        get
        {
            _workbook.ThrowIfDisposed();
            var af = _underlying.GetCTWorksheet().autoFilter;
            return af?.@ref;
        }
    }

    // ---- Per-column criteria (decision I-66) --------------------------

    public void SetAutoFilterColumn(int columnOffset, FilterCriteria criteria)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(criteria);

        var af = _underlying.GetCTWorksheet().autoFilter
            ?? throw new InvalidOperationException(
                "No AutoFilter is set on this sheet. Call SetAutoFilter(range) first.");

        EnsureColumnInRange(af, columnOffset);

        af.filterColumn ??= new List<CT_FilterColumn>();

        // Replace any existing entry for this column.
        for (int i = af.filterColumn.Count - 1; i >= 0; i--)
        {
            if (af.filterColumn[i].colId == (uint)columnOffset)
                af.filterColumn.RemoveAt(i);
        }

        var fc = new CT_FilterColumn
        {
            colId = (uint)columnOffset,
            customFilters = BuildCustomFilters(criteria),
        };
        af.filterColumn.Add(fc);
    }

    public void ClearAutoFilterColumn(int columnOffset)
    {
        _workbook.ThrowIfDisposed();
        var af = _underlying.GetCTWorksheet().autoFilter;
        if (af?.filterColumn is null) return;

        for (int i = af.filterColumn.Count - 1; i >= 0; i--)
        {
            if (af.filterColumn[i].colId == (uint)columnOffset)
                af.filterColumn.RemoveAt(i);
        }
        if (af.filterColumn.Count == 0)
        {
            af.filterColumn = null;
        }
    }

    private static void EnsureColumnInRange(CT_AutoFilter af, int columnOffset)
    {
        if (columnOffset < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(columnOffset), columnOffset, "columnOffset must be >= 0.");
        }
        // af.@ref is the A1 range — parse and check bounds.
        var (r1, c1, _, c2) = CellAddress.ParseRange(af.@ref);
        _ = r1;   // suppress unused-var
        int rangeWidth = c2 - c1 + 1;
        if (columnOffset >= rangeWidth)
        {
            throw new ArgumentOutOfRangeException(
                nameof(columnOffset), columnOffset,
                $"columnOffset must be < AutoFilter range width ({rangeWidth} for '{af.@ref}').");
        }
    }

    private static CT_CustomFilters BuildCustomFilters(FilterCriteria criteria)
    {
        var list = new List<CT_CustomFilter>(2)
        {
            new CT_CustomFilter
            {
                @operator = ToCTOp(criteria.Op1),
                val = criteria.Val1,
            }
        };
        if (criteria.Combine != FilterCriteria.Combinator.Single)
        {
            list.Add(new CT_CustomFilter
            {
                @operator = ToCTOp(criteria.Op2!.Value),
                val = criteria.Val2!,
            });
        }

        return new CT_CustomFilters
        {
            customFilter = list,
            and = criteria.Combine == FilterCriteria.Combinator.And,
        };
    }

    private static ST_FilterOperator ToCTOp(FilterCriteria.Op op) => op switch
    {
        FilterCriteria.Op.Equal => ST_FilterOperator.equal,
        FilterCriteria.Op.NotEqual => ST_FilterOperator.notEqual,
        FilterCriteria.Op.GreaterThan => ST_FilterOperator.greaterThan,
        FilterCriteria.Op.GreaterThanOrEqual => ST_FilterOperator.greaterThanOrEqual,
        FilterCriteria.Op.LessThan => ST_FilterOperator.lessThan,
        FilterCriteria.Op.LessThanOrEqual => ST_FilterOperator.lessThanOrEqual,
        _ => ST_FilterOperator.equal,
    };
}
