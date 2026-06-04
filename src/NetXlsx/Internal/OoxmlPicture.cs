// I-82 engine swap — Open XML SDK-backed IPicture (drawings slice, pictures).
//
// A snapshot of one anchored picture on the SDK engine. Created by
// OoxmlSheet.AddPicture and surfaced (freshly per call) by OoxmlSheet.Pictures.
// The anchor geometry mirrors the NPOI engine's XssfPicture contract exactly so
// the eventual cutover is behavior-preserving (decision I-81):
//
//   FromCell / ToCell  A1 corners. For a one-cell anchor (xdr:oneCellAnchor) there
//                      is no distinct end cell, so ToCell == FromCell and
//                      Dx2 == Dy2 == 0 — the rendered size lives in <xdr:ext>, which
//                      this read surface does not expose (NPOI's IPicture does not
//                      either). Two-cell anchors carry both corners + all four EMU
//                      offsets (lesson #5: each image keeps its own dx/dy; the end
//                      cell is exclusive, the same convention CellAddress round-trips).
//   Data               the raw embedded image bytes.
//
// The escape hatch (Underlying, #32 / I-82) hands out the live xdr:pic element
// the snapshot was created from / parsed out of.

using System;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace NetXlsx;

internal sealed class OoxmlPicture : IPicture
{
    private readonly OoxmlWorkbook _workbook;
    private readonly OoxmlSheet _sheet;
    private readonly ImageFormat _format;
    private readonly string _fromCell;
    private readonly string _toCell;
    private readonly int _dx1, _dy1, _dx2, _dy2;
    private readonly byte[] _data;
    private readonly XDR.Picture _pic;

    internal OoxmlPicture(
        OoxmlWorkbook workbook, OoxmlSheet sheet, ImageFormat format,
        string fromCell, string toCell,
        int dx1, int dy1, int dx2, int dy2, byte[] data, XDR.Picture pic)
    {
        _workbook = workbook;
        _sheet = sheet;
        _format = format;
        _fromCell = fromCell;
        _toCell = toCell;
        _dx1 = dx1; _dy1 = dy1; _dx2 = dx2; _dy2 = dy2;
        _data = data;
        _pic = pic;
    }

    public ISheet Sheet { get { _workbook.ThrowIfDisposed(); return _sheet; } }
    public ImageFormat Format { get { _workbook.ThrowIfDisposed(); return _format; } }
    public string FromCell { get { _workbook.ThrowIfDisposed(); return _fromCell; } }
    public string ToCell { get { _workbook.ThrowIfDisposed(); return _toCell; } }
    public int Dx1 { get { _workbook.ThrowIfDisposed(); return _dx1; } }
    public int Dy1 { get { _workbook.ThrowIfDisposed(); return _dy1; } }
    public int Dx2 { get { _workbook.ThrowIfDisposed(); return _dx2; } }
    public int Dy2 { get { _workbook.ThrowIfDisposed(); return _dy2; } }
    public byte[] Data { get { _workbook.ThrowIfDisposed(); return _data; } }

    // Escape hatch (#32 / I-82): the live xdr:pic element. Disposal first.
    public XDR.Picture Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _pic; }
    }
}
