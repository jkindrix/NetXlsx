// XssfWorkbook — named-style registry per decisions I-57 (v1.1) + I-67 (v1.3).
// Core class structure is in XssfWorkbook.cs.

using System;
using System.Collections.Generic;
using System.Reflection;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed partial class XssfWorkbook
{
    private Dictionary<string, CellStyle>? _namedStyles;
    private bool _namedStylesRehydrated;

    private Dictionary<string, CellStyle> NamedStyles
    {
        get
        {
            if (_namedStyles is null)
            {
                _namedStyles = new Dictionary<string, CellStyle>(StringComparer.OrdinalIgnoreCase);
                RehydrateNamedStylesFromOoxml();
            }
            return _namedStyles;
        }
    }

    public void RegisterStyle(string name, CellStyle style)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(style);
        if (name.Length == 0)
            throw new ArgumentException("Style name cannot be empty.", nameof(name));
        NamedStyles[name] = style;
        WriteNamedStyleToOoxml(name, style);
    }

    public CellStyle? GetRegisteredStyle(string name)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        // Force lazy rehydration of OOXML-persisted entries before lookup.
        return NamedStyles.TryGetValue(name, out var s) ? s : null;
    }

    public StylePoolDiagnostics GetStylePoolDiagnostics()
    {
        ThrowIfDisposed();
        var pool = StylePool;
        return new StylePoolDiagnostics(
            styleHits: pool.StyleHitCount,
            styleMisses: pool.StyleMissCount,
            fontHits: pool.FontHitCount,
            fontMisses: pool.FontMissCount,
            uniqueStyles: pool.UniqueStyles,
            uniqueFonts: pool.UniqueFonts);
    }

    public IReadOnlyCollection<string> RegisteredStyleNames
    {
        get
        {
            ThrowIfDisposed();
            // Force lazy rehydration so a freshly-opened workbook
            // surfaces its OOXML-persisted name list.
            return NamedStyles.Keys;
        }
    }

    /// <summary>
    /// Internal lookup throwing the canonical "no such name" error message.
    /// Shared by ICell.ApplyNamedStyle and IRange.ApplyNamedStyle.
    /// </summary>
    internal CellStyle ResolveNamedStyleOrThrow(string name)
    {
        var style = GetRegisteredStyle(name);
        if (style is null)
        {
            throw new ArgumentException(
                $"No style is registered under '{name}'. " +
                "Use IWorkbook.RegisterStyle before referencing the name.",
                nameof(name));
        }
        return style;
    }

    // ---- OOXML round-trip (decision I-67) -----------------------------
    //
    // v1.3 makes named styles produce real entries in OOXML's
    // `cellStyleXfs` + `cellStyles` tables so they survive a save/open
    // round-trip and appear in Excel's "Cell Styles" panel. The
    // in-process `_namedStyles` dictionary remains the source of truth
    // at runtime; OOXML serialization is a side effect of RegisterStyle
    // and a hydrate hook on Open.

    private static readonly MethodInfo s_getCTStylesheet = typeof(StylesTable).GetMethod(
        "GetCTStylesheet",
        BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public)
        ?? throw new InvalidOperationException(
            "NPOI 2.7.3 internal change: StylesTable.GetCTStylesheet() not accessible.");

    private CT_Stylesheet GetCTStylesheet()
    {
        var styles = _underlying.GetStylesSource();
        return (CT_Stylesheet)s_getCTStylesheet.Invoke(styles, null)!;
    }

    private void WriteNamedStyleToOoxml(string name, CellStyle style)
    {
        // Allocate the visual style via the existing style-pool dedup —
        // that gives us the cellXfs entry whose font / fill / border
        // indices we'll mirror into a fresh CT_Xf in cellStyleXfs.
        var npoiStyle = (XSSFCellStyle)StylePool.GetOrCreate(style);
        var sourceXf = npoiStyle.GetCoreXf();

        var styles = _underlying.GetStylesSource();
        var ctss = GetCTStylesheet();

        // Build a copy of the visual CT_Xf for the named-style table.
        // applyFont/Fill/Border are gated on the index being non-zero
        // (zero = the workbook default font/fill/border which doesn't
        // need explicit apply).
        var namedXf = new CT_Xf
        {
            numFmtId = sourceXf.numFmtId,
            numFmtIdSpecified = sourceXf.numFmtIdSpecified,
            fontId = sourceXf.fontId,
            fontIdSpecified = true,
            fillId = sourceXf.fillId,
            fillIdSpecified = true,
            borderId = sourceXf.borderId,
            borderIdSpecified = true,
            applyFont = sourceXf.fontId > 0,
            applyFill = sourceXf.fillId > 0,
            applyBorder = sourceXf.borderId > 0,
            applyNumberFormat = sourceXf.numFmtId > 0,
            applyAlignment = sourceXf.applyAlignment,
            alignment = sourceXf.alignment,
        };

        // Append to NPOI's internal styleXfs list (the one that wins on
        // save) and capture the new index.
        int newSize = NpoiInternals.PutCellStyleXf(styles, namedXf);
        uint xfId = (uint)(newSize - 1);

        // Now register the named-style entry. If a CT_CellStyle with
        // this name already exists, update it; otherwise append.
        ctss.cellStyles ??= new CT_CellStyles { cellStyle = new List<CT_CellStyle>() };
        ctss.cellStyles.cellStyle ??= new List<CT_CellStyle>();

        CT_CellStyle? existing = null;
        foreach (var cs in ctss.cellStyles.cellStyle)
        {
            if (string.Equals(cs.name, name, StringComparison.OrdinalIgnoreCase))
            {
                existing = cs;
                break;
            }
        }
        if (existing is not null)
        {
            existing.xfId = xfId;
        }
        else
        {
            ctss.cellStyles.cellStyle.Add(new CT_CellStyle
            {
                name = name,
                xfId = xfId,
            });
        }
        ctss.cellStyles.count = (uint)ctss.cellStyles.cellStyle.Count;
    }

    /// <summary>
    /// Walks the workbook's OOXML named-style table on first access and
    /// populates <see cref="_namedStyles"/>. Called from the
    /// <see cref="NamedStyles"/> property's lazy initializer so a
    /// freshly-opened workbook's <c>GetRegisteredStyle</c> /
    /// <c>RegisteredStyleNames</c> calls return real data.
    /// </summary>
    private void RehydrateNamedStylesFromOoxml()
    {
        if (_namedStylesRehydrated) return;
        _namedStylesRehydrated = true;

        var ctss = GetCTStylesheet();
        if (ctss.cellStyles?.cellStyle is not { } entries) return;

        var styles = _underlying.GetStylesSource();

        foreach (var cs in entries)
        {
            // Skip nameless / built-in "Normal" entries — only surface
            // user-registered names (the convention: workbooks
            // without explicit RegisterStyle calls have just the
            // built-in Normal style which we don't want to expose
            // as a "registered" name).
            if (string.IsNullOrEmpty(cs.name)) continue;
            if (string.Equals(cs.name, "Normal", StringComparison.OrdinalIgnoreCase)) continue;

            var xf = NpoiInternals.GetCellStyleXfAt(styles, (int)cs.xfId);
            if (xf is null) continue;

            var cellStyle = ReadCellStyleFromXf(styles, xf);
            // Don't call RegisterStyle (avoid re-writing the same OOXML
            // entry); poke the dictionary directly.
            _namedStyles![cs.name] = cellStyle;
        }
    }

    /// <summary>
    /// Parses a <c>CT_Xf</c> from the named-style table back to a
    /// <see cref="CellStyle"/> value, looking up font / fill / border
    /// via the workbook's resolved tables. Inverse of the
    /// allocation path in <see cref="WriteNamedStyleToOoxml"/>.
    /// </summary>
    private static CellStyle ReadCellStyleFromXf(StylesTable styles, CT_Xf xf)
    {
        // For most reads we can lean on the existing
        // CellStylePool.ReadFromNpoi by materializing the CT_Xf as a
        // temporary cellXfs entry and wrapping in XSSFCellStyle. The
        // duplicate cellXfs entry is one-per-named-style at read time
        // — acceptable given the rarity of named-style heavy workbooks.
        //
        // Build a fresh cellXfs entry that mirrors the named-style XF.
        var temp = new CT_Xf
        {
            numFmtId = xf.numFmtId, numFmtIdSpecified = xf.numFmtIdSpecified,
            fontId = xf.fontId, fontIdSpecified = true,
            fillId = xf.fillId, fillIdSpecified = true,
            borderId = xf.borderId, borderIdSpecified = true,
            applyFont = xf.fontId > 0,
            applyFill = xf.fillId > 0,
            applyBorder = xf.borderId > 0,
            applyNumberFormat = xf.numFmtId > 0,
            applyAlignment = xf.applyAlignment,
            alignment = xf.alignment,
        };

        // PutCellXf is internal on StylesTable; reach via NpoiInternals.
        int newSize = NpoiInternals.PutCellXf(styles, temp);
        // PutCellXf returns the new size; GetStyleAt takes 0-based index.
        var npoiStyle = styles.GetStyleAt(newSize - 1);

        return CellStylePool.ReadFromNpoi(npoiStyle);
    }
}
