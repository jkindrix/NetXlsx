// I-82 engine swap — Open XML SDK-backed IRow (cells & rows slice).
//
// A row handle over a 1-based row index. Cell access and value setters delegate
// to OoxmlCell. Row height / hidden are layout concerns and throw NotYet until
// their slice. The DateTime/DateOnly/TimeOnly/TimeSpan Set overloads delegate to
// the cell's (deferred) date/time setters, so they throw NotYet consistently.

using System;
using System.Runtime.CompilerServices;

namespace NetXlsx;

internal sealed class OoxmlRow : IRow
{
    private readonly OoxmlSheet _sheet;
    private readonly int _index;

    internal OoxmlRow(OoxmlSheet sheet, int index)
    {
        _sheet = sheet;
        _index = index;
    }

    public int Index { get { _sheet.WorkbookInternal.ThrowIfDisposed(); return _index; } }
    public ISheet Sheet { get { _sheet.WorkbookInternal.ThrowIfDisposed(); return _sheet; } }

    public ICell Cell(int column)
    {
        _sheet.WorkbookInternal.ThrowIfDisposed();
        if (column < 1 || column > CellAddress.MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(column), column, $"column must be in [1, {CellAddress.MaxColumn}]");
        return new OoxmlCell(_sheet, _index, column);
    }

    public ICell this[int column] => Cell(column);
    public ICell this[string columnLetter] => Cell(CellAddress.ParseColumn(columnLetter));

    public IRow Set(int column, string value) { Cell(column).SetString(value); return this; }
    public IRow Set(int column, double value) { Cell(column).SetNumber(value); return this; }
    public IRow Set(int column, decimal value) { Cell(column).SetNumber(value); return this; }
    public IRow Set(int column, int value) { Cell(column).SetNumber(value); return this; }
    public IRow Set(int column, long value) { Cell(column).SetNumber(value); return this; }
    public IRow Set(int column, bool value) { Cell(column).SetBool(value); return this; }
    public IRow Set(int column, DateTime value) { Cell(column).SetDate(value); return this; }
    public IRow Set(int column, DateOnly value) { Cell(column).SetDate(value); return this; }
    public IRow Set(int column, TimeOnly value) { Cell(column).SetTime(value); return this; }
    public IRow Set(int column, TimeSpan value) { Cell(column).SetDuration(value); return this; }

    // ---- Deferred (layout slice; see I-82) ---------------------------------

    private static NotImplementedException NotYet([CallerMemberName] string? member = null)
        => new(
            $"IRow.{member} is not yet implemented on the Open XML SDK engine " +
            "(I-82 engine swap). Row layout (height/hidden) lands in a later slice; " +
            "track the swap in docs/design.md (I-82).");

    public float HeightInPoints { get => throw NotYet(); set => throw NotYet(); }
    public bool Hidden { get => throw NotYet(); set => throw NotYet(); }

    public NPOI.XSSF.UserModel.XSSFRow Underlying => throw new NotSupportedException(
        "IRow.Underlying (NPOI XSSFRow) is not available on the Open XML SDK " +
        "engine (I-82). Use IWorkbook.OpenXmlDocument for the SDK escape hatch.");
}
