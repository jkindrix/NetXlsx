// XssfPicture — internal wrapper around NPOI's XSSFPicture.
// Created via XssfSheet.AddPicture; not public-constructible.

using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed class XssfPicture : IPicture
{
    private readonly XssfWorkbook _workbook;
    private readonly XssfSheet _sheet;
    private readonly XSSFPicture _underlying;
    private readonly ImageFormat _format;

    public XssfPicture(XssfWorkbook workbook, XssfSheet sheet, XSSFPicture underlying, ImageFormat format)
    {
        _workbook = workbook;
        _sheet = sheet;
        _underlying = underlying;
        _format = format;
    }

    public ISheet Sheet
    {
        get { _workbook.ThrowIfDisposed(); return _sheet; }
    }

    public ImageFormat Format
    {
        get { _workbook.ThrowIfDisposed(); return _format; }
    }

    public XSSFPicture Underlying
    {
        get { _workbook.ThrowIfDisposed(); return _underlying; }
    }
}
