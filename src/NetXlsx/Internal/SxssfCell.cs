using System;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;

namespace NetXlsx;

internal sealed class SxssfCell : IStreamingCell
{
    private readonly SxssfWorkbook _workbook;
    private readonly SXSSFCell _underlying;
    private readonly int _row1;
    private readonly int _col1;

    public SxssfCell(SxssfWorkbook workbook, SXSSFCell underlying, int row1, int col1)
    {
        _workbook = workbook;
        _underlying = underlying;
        _row1 = row1;
        _col1 = col1;
    }

    public string Address { get { _workbook.ThrowIfDisposed(); return CellAddress.Format(_row1, _col1); } }
    public int RowIndex { get { _workbook.ThrowIfDisposed(); return _row1; } }
    public int ColumnIndex { get { _workbook.ThrowIfDisposed(); return _col1; } }

    public CellKind Kind
    {
        get
        {
            _workbook.ThrowIfDisposed();
            return _underlying.CellType switch
            {
                CellType.Blank => CellKind.Empty,
                CellType.String => CellKind.String,
                CellType.Boolean => CellKind.Bool,
                CellType.Error => CellKind.Error,
                CellType.Formula => CellKind.Formula,
                CellType.Numeric => DateUtil.IsCellDateFormatted(_underlying) ? CellKind.Date : CellKind.Number,
                _ => CellKind.Empty,
            };
        }
    }

    public void SetString(string value)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(value);
        _underlying.SetCellValue(value);
    }

    public void SetNumber(double value) { _workbook.ThrowIfDisposed(); _underlying.SetCellValue(value); }
    public void SetNumber(decimal value) { _workbook.ThrowIfDisposed(); _underlying.SetCellValue((double)value); }
    public void SetNumber(int value) { _workbook.ThrowIfDisposed(); _underlying.SetCellValue((double)value); }
    public void SetNumber(long value) { _workbook.ThrowIfDisposed(); _underlying.SetCellValue((double)value); }
    public void SetBool(bool value) { _workbook.ThrowIfDisposed(); _underlying.SetCellValue(value); }

    public void SetDate(DateTime value)
    {
        _workbook.ThrowIfDisposed();
        _underlying.SetCellValue(value);
        ApplyDefaultStyleIfUnstyled(_workbook.StylePool.GetOrCreate(new CellStyle { NumberFormat = "yyyy-mm-dd hh:mm:ss" }));
    }

    public void SetFormula(string formula)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(formula);
        var body = formula.Length > 0 && formula[0] == '=' ? formula.Substring(1) : formula;
        if (body.Length == 0)
            throw new FormulaException("formula body is empty (expected '=...' or a non-empty expression)");
        try
        {
            _underlying.SetCellFormula(body);
        }
        catch (Exception ex)
        {
            throw new FormulaException($"failed to parse formula '{formula}': {ex.Message}", ex);
        }
    }

    public IStreamingCell Style(CellStyle style)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(style);
        var existingCellStyle = _underlying.CellStyle;
        // Merge non-null overlay axes over the cell's current style.
        var merged = MergeOverlayOverExisting(style, existingCellStyle);
        _underlying.CellStyle = _workbook.StylePool.GetOrCreate(merged);
        return this;
    }

    public IStreamingCell NumberFormat(string format)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(format);
        return Style(new CellStyle { NumberFormat = format });
    }

    private void ApplyDefaultStyleIfUnstyled(ICellStyle defaultStyle)
    {
        if (_underlying.CellStyle.Index == 0)
            _underlying.CellStyle = defaultStyle;
    }

    // Match XssfCell.Merge semantics — overlay-non-null wins, existing
    // axes are preserved for null overlay properties. For streaming
    // we read the few fields we model back off the existing NPOI
    // ICellStyle so callers can do incremental Style() calls.
    private static CellStyle MergeOverlayOverExisting(CellStyle overlay, ICellStyle existing)
    {
        // Conservative read-back: only NumberFormat is reliably
        // round-trippable through NPOI's ICellStyle for streaming cells;
        // other axes (font, fill, borders) live in style-table indices
        // we don't keep a reverse lookup for. Streaming style use is
        // typically one-shot per cell, so a null existing axis is fine.
        return new CellStyle
        {
            NumberFormat = overlay.NumberFormat ?? (existing?.GetDataFormatString()),
            Bold = overlay.Bold,
            Italic = overlay.Italic,
            Underline = overlay.Underline,
            FontName = overlay.FontName,
            FontSize = overlay.FontSize,
            FontColor = overlay.FontColor,
            Background = overlay.Background,
            HorizontalAlignment = overlay.HorizontalAlignment,
            VerticalAlignment = overlay.VerticalAlignment,
            WrapText = overlay.WrapText,
            Borders = overlay.Borders,
        };
    }
}
