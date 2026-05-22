// XssfCell — rich-text (multi-run formatted string) surface per
// decision I-50. Plain SetString → CellKind.String with
// NumFormattingRuns == 0; SetRichText → CellKind.String with one or
// more explicit formatting runs. GetRichText reads back only when
// formatting runs are present, distinguishing "user set rich text" /
// "file has rich text" from "plain string".

using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed partial class XssfCell
{
    public void SetRichText(RichText value)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(value);

        var plain = value.PlainText;
        int limit = _workbook.Options.MaxCellTextLength;
        if (plain.Length > limit)
        {
            throw new ResourceLimitExceededException(
                "cell text length", limit, plain.Length);
        }

        var rts = new XSSFRichTextString(plain);
        int cursor = 0;
        foreach (var run in value.Runs)
        {
            int runLen = run.Text.Length;
            if (runLen == 0) continue;  // empty runs contribute no formatting
            var font = _workbook.StylePool.GetOrCreateRunFont(run.Style);
            rts.ApplyFont(cursor, cursor + runLen, font);
            cursor += runLen;
        }
        _underlying.SetCellValue(rts);
    }

    public RichText? GetRichText()
    {
        _workbook.ThrowIfDisposed();
        if (_underlying.CellType != CellType.String) return null;

        var rts = _underlying.RichStringCellValue as XSSFRichTextString;
        if (rts is null) return null;
        if (rts.NumFormattingRuns <= 0) return null;

        var plain = rts.String ?? string.Empty;
        var runs = new List<RichTextRun>(rts.NumFormattingRuns);

        // NPOI exposes formatting-run boundaries via GetIndexOfFormattingRun(i).
        // Each run extends from its start index to the next run's start (or
        // the end of the string for the last run). Any prefix before the
        // first run's start is an unformatted run (default font).
        int firstStart = rts.GetIndexOfFormattingRun(0);
        if (firstStart > 0)
        {
            runs.Add(new RichTextRun(plain.Substring(0, firstStart), RichTextStyle.Default));
        }

        for (int i = 0; i < rts.NumFormattingRuns; i++)
        {
            int start = rts.GetIndexOfFormattingRun(i);
            int end = (i + 1 < rts.NumFormattingRuns)
                ? rts.GetIndexOfFormattingRun(i + 1)
                : plain.Length;
            if (end <= start) continue;
            var font = rts.GetFontOfFormattingRun(i);
            var style = font is null
                ? RichTextStyle.Default
                : CellStylePool.ReadRunStyleFromFont(font);
            runs.Add(new RichTextRun(plain.Substring(start, end - start), style));
        }

        return runs.Count == 0 ? null : new RichText(runs);
    }
}
