// I-82 engine swap — Open XML SDK-backed theme round-trip (drawings slice).
//
// A workbook's theme lives in xl/theme/theme1.xml as a dedicated ThemePart hung
// off the WorkbookPart by the "…/relationships/theme" relationship — NOT a child
// element of any container, so SDK-quirk #8 (schema-ordered insertion) does not
// apply here (the theme is a part, not a sibling in a strict CT sequence). The
// part graph is wired by the SDK: AddNewPart<ThemePart>() creates the part, its
// content type, and the relationship in one call; FeedData writes the raw bytes.
//
// Why preserve the theme at all (OOXML truth, lesson #2): a workbook with no
// theme part makes Excel fall back to built-in defaults that break column-width
// display and theme-color resolution. So the theme must round-trip faithfully.
// We write/read the bytes verbatim (never materializing ThemePart.Theme into a
// DOM that would re-serialize and drift), mirroring the NPOI engine which wrote
// the part bytes directly through the OPC package.
//
// The read-side resolution (ResolveThemeColor / GetThemeLineWidthEmu) is fully
// engine-agnostic: it delegates to ThemeInfo, the same parsed-bytes view the
// NPOI engine uses (decision I-81). Excel's tint algorithm, the OOXML cell-color
// slot mapping (0=lt1,1=dk1,2=lt2,3=dk2,4..9=accent1..6,10=hlink,11=folHlink),
// the tx1/bg1/tx2/bg2 aliases, and the indexed line-width table all live in
// ThemeInfo and are shared verbatim across both engines.

using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace NetXlsx;

internal sealed partial class OoxmlWorkbook
{
    // Parsed view of the workbook's theme1.xml, cached on first resolve call and
    // invalidated by SetThemeXml. Mirrors the NPOI engine's _themeInfo cache
    // (decision I-81) — the theme is constant between mutations, so one parse
    // serves every ResolveThemeColor / GetThemeLineWidthEmu call.
    private ThemeInfo? _themeInfo;

    public void SetThemeXml(byte[] themeXml)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(themeXml);

        var wbPart = _document.WorkbookPart
            ?? throw new InvalidOperationException("Workbook has no workbook part.");

        // AddNewPart<ThemePart>() also adds the content type + theme relationship;
        // reuse an existing part on a re-set so we don't orphan a relationship.
        var themePart = wbPart.ThemePart ?? wbPart.AddNewPart<ThemePart>();

        using (var ms = new MemoryStream(themeXml, writable: false))
            themePart.FeedData(ms);

        // Invalidate the cached resolve view so subsequent ResolveThemeColor calls
        // see the new theme (decision I-81).
        _themeInfo = null;
    }

    public byte[]? GetThemeXml()
    {
        ThrowIfDisposed();

        // Read the part stream directly (not ThemePart.Theme): the bytes are
        // authoritative and reflect a post-Open SetThemeXml without forcing a DOM
        // materialization that would later re-serialize on Save.
        var themePart = _document.WorkbookPart?.ThemePart;
        if (themePart is null) return null;

        using var stream = themePart.GetStream(FileMode.Open, FileAccess.Read);
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        var bytes = ms.ToArray();
        return bytes.Length > 0 ? bytes : null;
    }

    // Lazily parsed, then cached (invalidated by SetThemeXml). Identical shape to
    // the NPOI engine's Theme property.
    private ThemeInfo Theme => _themeInfo ??= ThemeInfo.Parse(GetThemeXml());

    public Color? ResolveThemeColor(int index, double tint = 0)
    {
        ThrowIfDisposed();
        return Theme.ResolveByIndex(index, tint);
    }

    public Color? ResolveThemeColor(ThemeColor color)
    {
        ThrowIfDisposed();
        return Theme.ResolveByIndex(color.Index, color.Tint);
    }

    public Color? ResolveThemeColor(string schemeName, double tint = 0)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(schemeName);
        return Theme.ResolveByName(schemeName, tint);
    }

    public int? GetThemeLineWidthEmu(int oneBasedIdx)
    {
        ThrowIfDisposed();
        return Theme.LineWidthEmu(oneBasedIdx);
    }
}
