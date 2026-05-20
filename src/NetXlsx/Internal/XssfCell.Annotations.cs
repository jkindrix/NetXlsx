// XssfCell — annotations. Comments (text + author per decision I11)
// and hyperlinks (scheme-sniffed per decision I13). Core class
// structure is in XssfCell.cs.

using System;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed partial class XssfCell
{
    public ICell Comment(string text, string? author = null)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(text);

        var factory = _workbook.Underlying.GetCreationHelper();
        var existing = _underlying.CellComment;
        if (existing is not null)
        {
            // NPOI rejects CreateCellComment on a cell that already has
            // one. Mutate in place instead.
            existing.String = factory.CreateRichTextString(text);
            existing.Author = author ?? "NetXlsx";
            return this;
        }

        var sheet = (XSSFSheet)_underlying.Sheet;
        var drawing = sheet.CreateDrawingPatriarch();

        var anchor = factory.CreateClientAnchor();
        anchor.Col1 = _col1 - 1;
        anchor.Row1 = _row1 - 1;
        anchor.Col2 = _col1 + 1;     // 2-column-wide popup
        anchor.Row2 = _row1 + 2;     // 2-row-tall popup

        var comment = drawing.CreateCellComment(anchor);
        comment.String = factory.CreateRichTextString(text);
        comment.Author = author ?? "NetXlsx";   // Decision I11.

        _underlying.CellComment = comment;
        return this;
    }

    public string? GetComment()
    {
        _workbook.ThrowIfDisposed();
        return _underlying.CellComment?.String?.String;
    }

    public string? GetCommentAuthor()
    {
        _workbook.ThrowIfDisposed();
        return _underlying.CellComment?.Author;
    }

    public ICell Hyperlink(string target, string? display = null)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(target);
        if (target.Length == 0)
            throw new ArgumentException("target cannot be empty", nameof(target));

        var (type, address) = SniffHyperlinkScheme(target);

        var link = new XSSFHyperlink(type)
        {
            Address = address,
        };
        if (display is not null) link.Label = display;
        _underlying.Hyperlink = link;

        // If a display value was supplied, set the cell's text to it
        // (replacing any prior value). If not and the cell is empty,
        // fall back to the raw target so the cell isn't blank.
        if (display is not null)
        {
            _underlying.SetCellValue(display);
        }
        else if (_underlying.CellType == NPOI.SS.UserModel.CellType.Blank)
        {
            _underlying.SetCellValue(target);
        }

        return this;
    }

    public string? GetHyperlink()
    {
        _workbook.ThrowIfDisposed();
        return _underlying.Hyperlink?.Address;
    }

    private static (NPOI.SS.UserModel.HyperlinkType type, string address) SniffHyperlinkScheme(string target)
    {
        // Decision I13: scheme-sniff into NPOI's HyperlinkType. We
        // preserve the original target text verbatim — the SchemeType
        // is what tells Excel/OPC how to interpret it.
        if (target.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || target.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            return (NPOI.SS.UserModel.HyperlinkType.Url, target);
        if (target.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
            return (NPOI.SS.UserModel.HyperlinkType.Email, target);
        if (target.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
            return (NPOI.SS.UserModel.HyperlinkType.File, target);
        if (target.StartsWith('#'))
            return (NPOI.SS.UserModel.HyperlinkType.Document, target.Substring(1));
        throw new ArgumentException(
            $"hyperlink target '{target}' uses an unsupported scheme. " +
            "Supported: http(s)://, mailto:, file://, internal #Sheet!Range " +
            "(decision I13).", nameof(target));
    }
}
