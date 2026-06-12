// I-82 engine swap — streaming cell (slice 9). A thin facade over the row's
// buffered cell data: setters mutate the in-window buffer; the OOXML <c> shape
// is materialized only when the row flushes (OoxmlStreamingRow.ToCell).
//
// Forward-only honesty: a write to a cell whose row has flushed past the
// window fails loud with InvalidOperationException. (NPOI's SXSSF keeps the
// orphaned row object in memory and silently discards such writes.) Same for
// writes after Save. Documented divergence — design.md I-82, streaming slice.

using System;

namespace NetXlsx;

internal sealed class OoxmlStreamingCell : IStreamingCell
{
    private readonly OoxmlStreamingWorkbook _workbook;
    private readonly StreamingRowBuffer _row;
    private readonly int _col1;

    internal OoxmlStreamingCell(OoxmlStreamingWorkbook workbook, StreamingRowBuffer row, int col1)
    {
        _workbook = workbook;
        _row = row;
        _col1 = col1;
    }

    public string Address
    {
        get { _workbook.ThrowIfDisposed(); return CellAddress.Format(_row.Row1, _col1); }
    }

    public int RowIndex
    {
        get { _workbook.ThrowIfDisposed(); return _row.Row1; }
    }

    public int ColumnIndex
    {
        get { _workbook.ThrowIfDisposed(); return _col1; }
    }

    public CellKind Kind
    {
        get
        {
            _workbook.ThrowIfDisposed();
            ThrowIfUnreachable();
            if (!_row.Cells.TryGetValue(_col1, out var data)) return CellKind.Empty;
            // A date is a number wearing a date format — same derivation as the
            // random-access SDK engine (OoxmlCell.Kind / pool.IsDateFormatted).
            if (data.Kind == CellKind.Number && _workbook.StylePool.IsDateFormatted(data.StyleIdx))
                return CellKind.Date;
            return data.Kind;
        }
    }

    // ---- Value setters ----------------------------------------------------------

    public void SetString(string value)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(value);
        int limit = _workbook.Options.MaxCellTextLength;
        if (value.Length > limit)
            throw new ResourceLimitExceededException("cell text length", limit, value.Length);
        var data = Write();
        data.Kind = CellKind.String;
        // ST_Xstring escaping (I-88) at the setter, same as the DOM engine.
        data.Str = XStringCodec.Encode(value);
        data.FormulaBody = null;
    }

    public void SetNumber(double value) => WriteNumber(value);

    // Decision I3.6 / §7.4: stored as IEEE-754 double; precision loss possible.
    public void SetNumber(decimal value) => WriteNumber((double)value);

    public void SetNumber(int value) => WriteNumber(value);

    // Values > 2^53 lose precision when stored as IEEE-754 double.
    public void SetNumber(long value) => WriteNumber(value);

    public void SetBool(bool value)
    {
        _workbook.ThrowIfDisposed();
        var data = Write();
        data.Kind = CellKind.Bool;
        data.Bool = value;
        data.Str = null;
        data.FormulaBody = null;
    }

    public void SetDate(DateTime value)
    {
        _workbook.ThrowIfDisposed();
        var data = Write();
        data.Kind = CellKind.Number;
        data.Num = OoxmlWorkbook.ToSerial(value, _workbook.Date1904);
        data.Str = null;
        data.FormulaBody = null;
        // Default datetime format only when the cell is unstyled — mirrors the
        // NPOI streaming engine's ApplyDefaultStyleIfUnstyled (and the DOM
        // engines' decision I-18: a user-set style is preserved).
        if (data.StyleIdx == 0)
        {
            data.StyleIdx = _workbook.StylePool.GetOrCreate(OoxmlWorkbook.DateTimeStyleSpec);
            data.AppliedStyle = OoxmlWorkbook.DateTimeStyleSpec;
        }
    }

    public void SetFormula(string formula)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(formula);
        // Same normalization + structural validation as the random-access SDK
        // engine (the SDK has no formula parser; semantic errors are Excel's
        // call — documented divergence, design.md I-82).
        var body = formula.Length > 0 && formula[0] == '=' ? formula.Substring(1) : formula;
        if (body.Length == 0)
            throw new FormulaException("formula body is empty (expected '=...' or a non-empty expression)");
        OoxmlCell.ValidateFormulaStructure(formula, body);
        var data = Write();
        data.Kind = CellKind.Formula;
        data.FormulaBody = body;
        data.Str = null;
    }

    // ---- Styling ----------------------------------------------------------------

    public IStreamingCell Style(CellStyle style)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(style);
        var data = Write();
        // Merge semantics mirror the NPOI streaming engine's
        // MergeOverlayOverExisting: only NumberFormat falls back to the cell's
        // current style; every other axis comes from the overlay verbatim.
        // (Streaming style use is typically one-shot per cell; the rich
        // read-merge-write of the random-access engines needs read-back the
        // forward-only window deliberately doesn't keep.)
        var merged = new CellStyle
        {
            NumberFormat = style.NumberFormat ?? data.AppliedStyle?.NumberFormat,
            Bold = style.Bold,
            Italic = style.Italic,
            Underline = style.Underline,
            FontName = style.FontName,
            FontSize = style.FontSize,
            FontColor = style.FontColor,
            Background = style.Background,
            BackgroundTheme = style.BackgroundTheme,
            HorizontalAlignment = style.HorizontalAlignment,
            VerticalAlignment = style.VerticalAlignment,
            WrapText = style.WrapText,
            Borders = style.Borders,
        };
        data.StyleIdx = _workbook.StylePool.GetOrCreate(merged);
        data.AppliedStyle = merged;
        return this;
    }

    public IStreamingCell NumberFormat(string format)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(format);
        return Style(new CellStyle { NumberFormat = format });
    }

    // ---- Buffer access ------------------------------------------------------------

    private void WriteNumber(double value)
    {
        _workbook.ThrowIfDisposed();
        var data = Write();
        data.Kind = CellKind.Number;
        data.Num = value;
        data.Str = null;
        data.FormulaBody = null;
    }

    private StreamingCellData Write()
    {
        ThrowIfUnreachable();
        if (!_row.Cells.TryGetValue(_col1, out var data))
        {
            data = new StreamingCellData();
            _row.Cells.Add(_col1, data);
        }
        return data;
    }

    private void ThrowIfUnreachable()
    {
        _workbook.ThrowIfSaved();
        if (_row.Flushed)
            throw new InvalidOperationException(
                $"row {_row.Row1} has been flushed past the row-access window and is on disk. " +
                "Streaming cells are forward-only: once a row flushes it cannot be read or " +
                "written. Raise StreamingOptions.RowAccessWindowSize to keep more rows " +
                "accessible (design.md I-82, streaming slice).");
    }
}
