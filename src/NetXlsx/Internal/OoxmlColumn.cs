// I-82 engine swap — Open XML SDK-backed IColumn (cells & rows slice).
//
// A lightweight column handle: Index / Letter / Sheet are live. Width, hidden,
// AutoSize, default style, and populated-cell iteration are column-formatting /
// layout concerns that land in later slices and throw NotYet for now.

using System;
using System.Runtime.CompilerServices;

namespace NetXlsx;

internal sealed class OoxmlColumn : IColumn
{
    private readonly OoxmlSheet _sheet;
    private readonly int _index;

    internal OoxmlColumn(OoxmlSheet sheet, int index)
    {
        _sheet = sheet;
        _index = index;
    }

    public int Index { get { _sheet.WorkbookInternal.ThrowIfDisposed(); return _index; } }
    public string Letter { get { _sheet.WorkbookInternal.ThrowIfDisposed(); return CellAddress.FormatColumn(_index); } }
    public ISheet Sheet { get { _sheet.WorkbookInternal.ThrowIfDisposed(); return _sheet; } }

    // ---- Deferred (styles / layout slices; see I-82) -----------------------

    private static NotImplementedException NotYet([CallerMemberName] string? member = null)
        => new(
            $"IColumn.{member} is not yet implemented on the Open XML SDK engine " +
            "(I-82 engine swap). Column width/hidden/style land in a later slice; " +
            "track the swap in docs/design.md (I-82).");

    public bool Hidden { get => throw NotYet(); set => throw NotYet(); }
    public double WidthUnits { get => throw NotYet(); set => throw NotYet(); }
    public IColumn Width(double units) => throw NotYet();
    public IColumn AutoSize() => throw NotYet();
    public IColumn ForEachPopulated(Action<ICell> apply) => throw NotYet();
    public IColumn SetDefaultStyle(CellStyle style) => throw NotYet();
}
