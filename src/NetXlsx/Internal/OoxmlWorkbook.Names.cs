// I-82 engine swap — Open XML SDK-backed named ranges (structure slice).
//
// OOXML stores named ranges as a <definedNames> container in workbook.xml, one
// <definedName name="…" [localSheetId="…"]>refers-to</definedName> per name.
// In CT_Workbook's strict child sequence <definedNames> sits between <sheets>
// and <calcPr> (SDK-quirk #3 — AppendChild does not reorder), so it is inserted
// immediately after <sheets>. The schema-validation gate (slice 4b) holds this
// to OpenXmlValidator on every named-range fixture.
//
// Contract mirrors the NPOI engine (XssfWorkbook.AddNamedRange): null/empty
// guards on name + formula; an optional sheetScope is resolved to a 0-based
// localSheetId or throws SheetNameException; names must be unique workbook-wide
// (case-insensitive) regardless of scope; a leading '=' on the formula is
// stripped. Excel name rules (start with letter/underscore; letters/digits/
// underscore/period; not a cell reference) are enforced here — the SDK has no
// built-in name validator, so this reproduces the documented IWorkbook contract
// that the NPOI engine delegated to NPOI's XSSFName.ValidateName.

using System;
using System.Collections.Generic;
using System.Linq;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlWorkbook
{
    public INamedRange AddNamedRange(string name, string formula, string? sheetScope = null)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(formula);
        if (name.Length == 0)
            throw new ArgumentException("name cannot be empty", nameof(name));
        if (formula.Length == 0)
            throw new ArgumentException("formula cannot be empty", nameof(formula));

        ValidateDefinedName(name);

        uint? localSheetId = null;
        if (sheetScope is not null)
        {
            int idx = IndexOfSheet(sheetScope);
            if (idx < 0)
                throw new SheetNameException(sheetScope, "no sheet with that name exists in this workbook (sheetScope)");
            localSheetId = (uint)idx;
        }

        // Excel permits a workbook-scope and a same-text sheet-scope name to
        // coexist, but the NPOI engine rejects this (XSSFName.ValidateName); v1
        // requires names unique workbook-wide regardless of scope. The SDK engine
        // matches that contract so behavior is identical across engines.
        var container = DefinedNamesContainer();
        if (container is not null)
        {
            foreach (var existing in container.Elements<S.DefinedName>())
                if (string.Equals(existing.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException(
                        $"a named range '{name}' already exists in the workbook " +
                        "(case-insensitive). Names must be unique workbook-wide " +
                        "regardless of scope.", nameof(name));
        }

        // Strip an optional leading '=' for consistency with SetFormula / the NPOI engine.
        var body = formula[0] == '=' ? formula.Substring(1) : formula;

        var defined = new S.DefinedName { Name = name, Text = body };
        if (localSheetId is uint lsid) defined.LocalSheetId = lsid;

        GetOrCreateDefinedNames().AppendChild(defined);
        return new OoxmlNamedRange(this, defined);
    }

    public IReadOnlyList<INamedRange> NamedRanges
    {
        get
        {
            ThrowIfDisposed();
            var container = DefinedNamesContainer();
            if (container is null) return Array.Empty<INamedRange>();
            var list = new List<INamedRange>();
            foreach (var defined in container.Elements<S.DefinedName>())
                list.Add(new OoxmlNamedRange(this, defined));
            return list.Count == 0 ? Array.Empty<INamedRange>() : list;
        }
    }

    private S.DefinedNames? DefinedNamesContainer()
        => _document.WorkbookPart?.Workbook?.GetFirstChild<S.DefinedNames>();

    // <definedNames> sits between <sheets> and <calcPr> in CT_Workbook.
    private S.DefinedNames GetOrCreateDefinedNames()
    {
        var workbook = _document.WorkbookPart!.Workbook!;
        var existing = workbook.GetFirstChild<S.DefinedNames>();
        if (existing is not null) return existing;

        var dn = new S.DefinedNames();
        var sheets = workbook.GetFirstChild<S.Sheets>();
        if (sheets is not null)
        {
            workbook.InsertAfter(dn, sheets);
        }
        else
        {
            // <sheets> is always present on a created/opened workbook; this branch
            // is defensive. Fall back to before <calcPr>, else append.
            var calcPr = workbook.GetFirstChild<S.CalculationProperties>();
            if (calcPr is not null) workbook.InsertBefore(dn, calcPr);
            else workbook.AppendChild(dn);
        }
        return dn;
    }

    // Excel defined-name rules (the documented IWorkbook.AddNamedRange contract):
    // 1-255 chars; first char a letter or underscore; remaining letters / digits /
    // underscore / period; must not look like a cell reference (e.g. A1, AB12).
    private static void ValidateDefinedName(string name)
    {
        if (name.Length > 255)
            throw new ArgumentException(
                $"named range name '{name}' exceeds Excel's 255-character limit", nameof(name));

        char first = name[0];
        if (!(char.IsLetter(first) || first == '_'))
            throw new ArgumentException(
                $"named range name '{name}' must start with a letter or underscore", nameof(name));

        foreach (char c in name)
            if (!(char.IsLetterOrDigit(c) || c == '_' || c == '.'))
                throw new ArgumentException(
                    $"named range name '{name}' may contain only letters, digits, " +
                    "underscores, and periods", nameof(name));

        if (LooksLikeCellReference(name))
            throw new ArgumentException(
                $"named range name '{name}' collides with a cell reference", nameof(name));
    }

    // True for an A1-style reference: 1-3 leading letters followed by digits and
    // nothing else (A1, AB12, XFD1048576). Catches the common collision the
    // contract calls out; deliberately conservative.
    private static bool LooksLikeCellReference(string name)
    {
        int i = 0;
        while (i < name.Length && IsAsciiLetter(name[i])) i++;
        if (i == 0 || i > 3 || i == name.Length) return false;
        for (int j = i; j < name.Length; j++)
            if (name[j] is < '0' or > '9') return false;
        return true;
    }

    private static bool IsAsciiLetter(char c) => c is (>= 'A' and <= 'Z') or (>= 'a' and <= 'z');
}
