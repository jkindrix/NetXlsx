// XssfCell — styling. Style merge semantics + NumberFormat shortcut +
// GetStyle reader + the private ReadCurrentStyle / Merge helpers.
// Core class structure is in XssfCell.cs.

using System;

namespace NetXlsx;

internal sealed partial class XssfCell
{
    public ICell Style(CellStyle style)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(style);

        // Merge: non-null properties in `style` overwrite the cell's
        // current style on that axis; null means "inherit existing."
        // CellStyle.Default (all-null) is a no-op.
        var current = ReadCurrentStyle();
        var merged = Merge(current, style);

        var npoiStyle = _workbook.StylePool.GetOrCreate(merged);
        _underlying.CellStyle = npoiStyle;
        return this;
    }

    public ICell NumberFormat(string format)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(format);
        return Style(new CellStyle { NumberFormat = format });
    }

    public ICell ApplyNamedStyle(string name)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        return Style(_workbook.ResolveNamedStyleOrThrow(name));
    }

    public CellStyle GetStyle()
    {
        _workbook.ThrowIfDisposed();
        return ReadCurrentStyle();
    }

    private CellStyle ReadCurrentStyle()
    {
        var ns = _underlying.CellStyle;
        // Index 0 is the workbook default — treat as "no style applied".
        if (ns is null || ns.Index == 0) return CellStyle.Default;
        return CellStylePool.ReadFromNpoi(ns);
    }

    private static CellStyle Merge(CellStyle existing, CellStyle overlay) => new()
    {
        Bold = overlay.Bold ?? existing.Bold,
        Italic = overlay.Italic ?? existing.Italic,
        Underline = overlay.Underline ?? existing.Underline,
        FontName = overlay.FontName ?? existing.FontName,
        FontSize = overlay.FontSize ?? existing.FontSize,
        FontColor = overlay.FontColor ?? existing.FontColor,
        Background = overlay.Background ?? existing.Background,
        NumberFormat = overlay.NumberFormat ?? existing.NumberFormat,
        HorizontalAlignment = overlay.HorizontalAlignment ?? existing.HorizontalAlignment,
        VerticalAlignment = overlay.VerticalAlignment ?? existing.VerticalAlignment,
        WrapText = overlay.WrapText ?? existing.WrapText,
        Borders = overlay.Borders ?? existing.Borders,
    };
}
