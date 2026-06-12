// I-82 engine swap — Open XML SDK-backed streaming writer (slice 9, the SXSSF
// replacement behind Workbook.CreateStreaming).
//
// Architecture (probed in /tmp/reflcon before implementation):
//   - Each sheet streams its worksheet XML into its OWN temp file through an
//     OpenXmlPartWriter the moment rows flush — the SXSSF temp-file model. This
//     honors the SDK writer's forward-only shape without faking random access,
//     keeps memory bounded by StreamingOptions.RowAccessWindowSize, and lets
//     callers interleave writes across sheets (one part stream per sheet, no
//     concurrent-package-stream constraint).
//   - Save (single-shot) finalizes every sheet's XML, then assembles the
//     package: SpreadsheetDocument.Create + workbook DOM + the detached style
//     pool's stylesheet + WorksheetPart.FeedData(tempFile) per sheet. The
//     worksheet bytes stream from temp file to package — never a whole-sheet
//     DOM.
//   - Styles dedup through OoxmlStylePool.CreateDetached: cellXfs indices are
//     allocated live while rows are written; the stylesheet attaches to the
//     WorkbookStylesPart at assembly.
//
// Contract honesty (vs the NPOI SXSSF engine, probed 2026-06-03):
//   - Save is SINGLE-SHOT on both engines. NPOI leaks ObjectDisposedException
//     ("Cannot write to a closed TextWriter") from SheetDataWriter internals on
//     a second Save; this engine throws a deliberate InvalidOperationException
//     up front. Same fail-loud contract, honest message.
//   - NPOI silently ACCEPTS AppendRow / cell writes after Save — the data goes
//     nowhere (the writer is closed; rows are lost). This engine fails loud
//     with InvalidOperationException, per the I-83 honesty discipline. One-sided
//     strictness divergence, documented in design.md.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed class OoxmlStreamingWorkbook : IStreamingWorkbook
{
    private readonly StreamingOptions _options;
    private readonly List<OoxmlStreamingSheet> _sheets = new();
    private readonly Dictionary<string, OoxmlStreamingSheet> _sheetsByName = new(StringComparer.OrdinalIgnoreCase);
    private readonly OoxmlStylePool _stylePool;
    private bool _saved;
    private bool _disposed;

    internal OoxmlStreamingWorkbook(StreamingOptions options)
    {
        _options = options;
        // Streaming workbooks are always newly created (write-only), so the
        // default-font / 1904-epoch options apply unconditionally (design #15),
        // exactly like the NPOI streaming engine.
        _stylePool = OoxmlStylePool.CreateDetached(options);
    }

    internal OoxmlStylePool StylePool => _stylePool;
    internal StreamingOptions Options => _options;
    internal bool Date1904 => _options.DateSystem == DateSystem.Excel1904;
    internal bool Saved => _saved;

    public IStreamingSheet AddSheet(string name)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        ThrowIfSaved();
        Workbook.ValidateSheetName(name);
        if (_sheetsByName.ContainsKey(name))
            throw new SheetNameException(name, "a sheet with this name already exists (case-insensitive)");

        var sheet = new OoxmlStreamingSheet(this, name);
        _sheets.Add(sheet);
        _sheetsByName[name] = sheet;
        return sheet;
    }

    public void Save(string path)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(path);
        ThrowIfSaved();
        // R-2: sibling-temp + rename — a failure must never leave the
        // destination truncated. The temp opens before the callback runs and
        // its name embeds the destination filename, preserving the original
        // guarantee here: a bad path (missing dir, no permission, illegal
        // filename) still fails BEFORE the single-shot finalize burns the
        // workbook's only Save.
        AtomicFileWriter.Write(path, fs =>
        {
            FinalizeSheets();
            AssembleTo(fs);
        });
    }

    public void Save(Stream stream, bool leaveOpen = true)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));
        ThrowIfSaved();

        // Assemble into a temp file, then copy the finished bytes out. The
        // package writer needs a seekable target; the caller's stream need only
        // be writable (matching the NPOI streaming engine's contract). Memory
        // stays bounded — the copy is a stream copy, not a buffer.
        var tmp = Path.Combine(Path.GetTempPath(), $"netxlsx-stream-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var fs = File.Create(tmp))
            {
                FinalizeSheets();
                AssembleTo(fs);
            }
            using (var src = File.OpenRead(tmp))
                src.CopyTo(stream);
        }
        finally
        {
            try { File.Delete(tmp); }
            catch (IOException) { /* best-effort temp cleanup */ }
        }
        if (!leaveOpen) stream.Dispose();
    }

    public Task SaveAsync(string path, CancellationToken ct = default)
    {
        ThrowIfDisposed();
        ct.ThrowIfCancellationRequested();
        return Task.Run(() => Save(path), ct);
    }

    public Task SaveAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default)
    {
        ThrowIfDisposed();
        ct.ThrowIfCancellationRequested();
        return Task.Run(() => Save(stream, leaveOpen), ct);
    }

    // No escape hatch on the streaming engine (v2.0.0 / I-82): rows stream
    // forward-only through OpenXmlWriter into per-sheet temp streams and the
    // package is assembled only at Save — there is no live document to expose.

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        // Best-effort per sheet: close any open writer/stream and delete the
        // temp file. A failure on one sheet must not leak the others' files.
        foreach (var sheet in _sheets)
            sheet.DisposeTemp();
    }

    public ValueTask DisposeAsync()
    {
        Dispose();
        return ValueTask.CompletedTask;
    }

    // ---- Save-time assembly --------------------------------------------------

    private void FinalizeSheets()
    {
        _saved = true;
        foreach (var sheet in _sheets)
            sheet.FinalizeXml();
    }

    private void AssembleTo(Stream target)
    {
        // autoSave:false + the explicit Save() below (R-35, streaming flavor):
        // with autosave on, a mid-assembly exception would re-serialize the
        // half-built package inside the using-dispose during unwind and mask
        // the original exception with the SDK's.
        using var doc = SpreadsheetDocument.Create(target, SpreadsheetDocumentType.Workbook, autoSave: false);
        var wbPart = doc.AddWorkbookPart();

        var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = _stylePool.Stylesheet;

        var sheets = new S.Sheets();
        uint sheetId = 1;
        foreach (var sheet in _sheets)
        {
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            // Dimension-splicing copy (R-13) — see WriteFinalizedXml.
            using (var dst = wsPart.GetStream(FileMode.Create, FileAccess.Write))
                sheet.WriteFinalizedXml(dst);
            sheets.AppendChild(new S.Sheet
            {
                Name = sheet.SheetName,
                SheetId = sheetId++,
                Id = wbPart.GetIdOfPart(wsPart),
            });
        }

        // CT_Workbook schema order: workbookPr precedes sheets (lesson #9 — only
        // Create applies the option; a streaming workbook is always created).
        var wb = new S.Workbook();
        if (Date1904)
            wb.AppendChild(new S.WorkbookProperties { Date1904 = true });
        wb.AppendChild(sheets);
        wbPart.Workbook = wb;

        // Flush the DOM-rooted parts (workbook, stylesheet) explicitly — the
        // FeedData'd worksheet parts are already raw bytes in the package.
        // With autoSave:false this is the only serialization point.
        doc.Save();
    }

    // ---- Guards ---------------------------------------------------------------

    internal void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }

    internal void ThrowIfSaved()
    {
        if (_saved)
            throw new InvalidOperationException(
                "The streaming workbook has already been saved. Streaming Save is " +
                "single-shot: rows stream forward-only to disk and the worksheet XML is " +
                "finalized at Save, so the workbook cannot be appended to or saved again. " +
                "(The NPOI engine silently loses rows appended after Save; this engine " +
                "fails loud — see design.md I-82, streaming slice.)");
    }
}
