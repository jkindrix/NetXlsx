// Foreign (non-Ooxml*) implementations of the drawing handle interfaces, used
// to exercise two paths in the I-91 slice-2 removal family:
//   1. ISheet.RemovePicture/RemoveChart/RemoveShape/RemoveConnector reject a
//      handle that is not one of ours with ArgumentException (the RemoveTable
//      foreign-handle contract).
//   2. The disposal matrix passes them after dispose, where the sheet's
//      ThrowIfUnusable fires first (ObjectDisposedException) before the
//      foreign-rejection ever runs — the same reason RemoveTable's row uses a
//      stub (see XssfTableStub).
// Every member is a benign default; these are never dereferenced past the
// type check.

using System;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace NetXlsx.Tests;

internal sealed class ForeignPicture : IPicture
{
    public ISheet Sheet => null!;
    public ImageFormat Format => ImageFormat.Png;
    public string FromCell => "";
    public string ToCell => "";
    public int Dx1 => 0;
    public int Dy1 => 0;
    public int Dx2 => 0;
    public int Dy2 => 0;
    public byte[] Data => Array.Empty<byte>();
    public PictureBorder? Border { get => null; set { } }
    public XDR.Picture Underlying => null!;
}

internal sealed class ForeignChart : IChart
{
    public ISheet Sheet => null!;
    public ChartType Type => ChartType.Line;
    public void SetTitle(string title) { }
    public DocumentFormat.OpenXml.Packaging.ChartPart Underlying => null!;
}

internal sealed class ForeignShape : IShape
{
    public ISheet Sheet => null!;
    public ShapeType Type => ShapeType.Rectangle;
    public XDR.Shape Underlying => null!;
}

internal sealed class ForeignConnector : IConnector
{
    public ISheet Sheet => null!;
    public ConnectorType Type => ConnectorType.Straight;
    public string FromCell => "";
    public string ToCell => "";
    public int Dx1 => 0;
    public int Dy1 => 0;
    public int Dx2 => 0;
    public int Dy2 => 0;
    public bool FlipH => false;
    public bool FlipV => false;
    public ConnectorEnd HeadEnd => ConnectorEnd.None;
    public ConnectorEnd TailEnd => ConnectorEnd.None;
    public Color? LineColor => null;
    public string? LineSchemeColor => null;
    public double? LineWidthPoints => null;
    public int? LineStyleRefIndex => null;
    public XDR.ConnectionShape Underlying => null!;
}
