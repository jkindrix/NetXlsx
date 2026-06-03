// I-82 engine swap — Open XML SDK-backed Excel Tables (CF/validation/tables/
// autofilter/sort slice). Surface per decision I-51.
//
// A table is a TableDefinitionPart hung off the WorksheetPart, referenced by
// a <tablePart r:id> child of the worksheet's <tableParts> container.
// <tableParts> is a 0..1 singleton near the end of CT_Worksheet's strict
// child sequence (routed through OoxmlSchemaOrder, SDK-quirk #8); its
// <tablePart> children are 0..*. Removing the last table drops the container
// (SDK-quirk #7 / NPOI parity — Excel is friendlier with it absent).
//
// Contract mirrors the NPOI engine (XssfSheet.Tables.cs): Excel name rules,
// workbook-wide name uniqueness across tables AND named ranges, the
// single-row-range / empty-header / duplicate-header rejections, and the
// belongs-to check on RemoveTable. Oracle divergence (vs the NPOI dump):
// this engine EMITS the schema-required <table @id> — NPOI 2.7.3 omits it,
// which OpenXmlValidator flags — and keeps <tableParts @count> in sync.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    public ITable AddTable(string a1Range, string name, string? style = null)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(a1Range);
        ArgumentNullException.ThrowIfNull(name);

        if (name.Length == 0)
            throw new ArgumentException("name cannot be empty.", nameof(name));
        ValidateTableName(name);
        EnsureNameIsUniqueWorkbookWide(name);

        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);
        if (r1 == r2)
        {
            throw new ArgumentException(
                "Table range must include at least one data row below the header " +
                $"(got '{a1Range}', a single-row range).", nameof(a1Range));
        }

        // Collect header cell values from the first row of the range — these
        // become the table's column names. Excel requires a non-empty string
        // in each header cell; reject blanks loud.
        var headers = new List<string>(c2 - c1 + 1);
        for (int col = c1; col <= c2; col++)
        {
            var cell = CellHandle(r1, col);
            var text = cell.Kind == CellKind.String ? cell.GetString() : null;
            if (string.IsNullOrEmpty(text))
            {
                throw new ArgumentException(
                    $"Header cell at {CellAddress.Format(r1, col)} is empty or non-string. " +
                    "Set header values on every column before calling AddTable.",
                    nameof(a1Range));
            }
            for (int j = 0; j < headers.Count; j++)
            {
                if (string.Equals(headers[j], text, StringComparison.OrdinalIgnoreCase))
                {
                    throw new ArgumentException(
                        $"Duplicate header '{text}' at {CellAddress.Format(r1, col)} — " +
                        "column names within a table must be unique (case-insensitive).",
                        nameof(a1Range));
                }
            }
            headers.Add(text);
        }

        // Build the table part. CT_Table children are a sequence:
        // autoFilter? sortState? tableColumns tableStyleInfo? extLst.
        var tablePart = _worksheetPart.AddNewPart<TableDefinitionPart>();
        var columns = new S.TableColumns { Count = (uint)headers.Count };
        for (int i = 0; i < headers.Count; i++)
            columns.AppendChild(new S.TableColumn { Id = (uint)(i + 1), Name = headers[i] });

        var table = new S.Table(columns)
        {
            Id = NextTableId(),
            Name = name,
            DisplayName = name,
            Reference = CellAddress.FormatRange(r1, c1, r2, c2),
        };
        if (style is not null)
            table.TableStyleInfo = new S.TableStyleInfo { Name = style };
        tablePart.Table = table;

        // Reference the part from the worksheet's <tableParts> container.
        var parts = OoxmlSchemaOrder.GetOrInsert(Worksheet, static () => new S.TableParts());
        parts.AppendChild(new S.TablePart { Id = _worksheetPart.GetIdOfPart(tablePart) });
        parts.Count = (uint)parts.Elements<S.TablePart>().Count();

        return new OoxmlTable(_workbook, this, tablePart);
    }

    public IReadOnlyList<ITable> Tables
    {
        get
        {
            _workbook.ThrowIfDisposed();
            // Enumerate via the <tableParts> children (document order) rather
            // than the part collection (arbitrary order).
            var container = Worksheet.GetFirstChild<S.TableParts>();
            if (container is null) return Array.Empty<ITable>();
            var list = new List<ITable>();
            foreach (var tp in container.Elements<S.TablePart>())
            {
                if (tp.Id?.Value is { } relId
                    && _worksheetPart.TryGetPartById(relId, out var part)
                    && part is TableDefinitionPart tdp)
                {
                    list.Add(new OoxmlTable(_workbook, this, tdp));
                }
            }
            return list.Count == 0 ? Array.Empty<ITable>() : list;
        }
    }

    public bool TryGetTable(string name, [MaybeNullWhen(false)] out ITable table)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        foreach (var t in Tables)
        {
            if (string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                table = t;
                return true;
            }
        }
        table = null;
        return false;
    }

    public void RemoveTable(ITable table)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(table);
        if (table is not OoxmlTable ot)
        {
            throw new ArgumentException(
                "Table instance is not an OoxmlTable — likely a foreign or mocked implementation.",
                nameof(table));
        }
        // Belongs-to check (and double-remove guard): the part must still be
        // a relation of this worksheet.
        if (!_worksheetPart.TableDefinitionParts.Contains(ot.Part))
        {
            throw new ArgumentException(
                "Table does not belong to this sheet (no relation found).",
                nameof(table));
        }
        string relId = _worksheetPart.GetIdOfPart(ot.Part);

        // Drop the <tablePart> entry; drop the whole container when empty.
        var container = Worksheet.GetFirstChild<S.TableParts>();
        if (container is not null)
        {
            foreach (var tp in container.Elements<S.TablePart>()
                         .Where(tp => tp.Id?.Value == relId).ToList())
                tp.Remove();
            int remaining = container.Elements<S.TablePart>().Count();
            if (remaining == 0) container.Remove();
            else container.Count = (uint)remaining;
        }

        // Drop the package relationship + part.
        _worksheetPart.DeletePart(ot.Part);
    }

    // ---- Internal helpers ---------------------------------------------

    // <table @id> must be unique workbook-wide; continue after the highest
    // id already present (an opened file may carry tables on any sheet).
    private uint NextTableId()
    {
        uint max = 0;
        var wbPart = _workbook.OpenXmlDocument!.WorkbookPart!;
        foreach (var ws in wbPart.WorksheetParts)
            foreach (var tdp in ws.TableDefinitionParts)
                if (tdp.Table?.Id?.Value is uint id && id > max) max = id;
        return max + 1;
    }

    /// <summary>
    /// Excel name rules per decision I-51 (mirrors named-range rules):
    /// must start with letter or underscore; subsequent chars letters /
    /// digits / underscores / periods. Must not look like a cell reference
    /// (e.g. "A1"). Spaces forbidden.
    /// </summary>
    private static void ValidateTableName(string name)
    {
        if (!(char.IsLetter(name[0]) || name[0] == '_'))
        {
            throw new ArgumentException(
                $"Table name '{name}' must start with a letter or underscore.",
                nameof(name));
        }
        for (int i = 1; i < name.Length; i++)
        {
            char c = name[i];
            if (!(char.IsLetterOrDigit(c) || c == '_' || c == '.'))
            {
                throw new ArgumentException(
                    $"Table name '{name}' contains invalid character '{c}' at index {i}. " +
                    "Allowed: letters, digits, underscore, period.",
                    nameof(name));
            }
        }
        if (LooksLikeCellAddress(name))
        {
            throw new ArgumentException(
                $"Table name '{name}' collides with an A1-style cell address. " +
                "Pick a name that cannot be confused with a cell reference.",
                nameof(name));
        }
    }

    private static bool LooksLikeCellAddress(string s)
    {
        int i = 0;
        while (i < s.Length && char.IsLetter(s[i])) i++;
        if (i == 0 || i == s.Length) return false;
        int letterEnd = i;
        while (i < s.Length && char.IsDigit(s[i])) i++;
        return i == s.Length && letterEnd > 0 && letterEnd < s.Length;
    }

    private void EnsureNameIsUniqueWorkbookWide(string name)
    {
        // Per Excel: table names share the namespace with named ranges and
        // with other tables in the workbook. Check both.
        foreach (var nr in _workbook.NamedRanges)
        {
            if (string.Equals(nr.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(
                    $"Name '{name}' is already taken by a named range " +
                    "(table and named-range names share the workbook-wide namespace).",
                    nameof(name));
            }
        }
        var wbPart = _workbook.OpenXmlDocument!.WorkbookPart!;
        foreach (var ws in wbPart.WorksheetParts)
        {
            foreach (var tdp in ws.TableDefinitionParts)
            {
                if (string.Equals(tdp.Table?.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
                {
                    throw new ArgumentException(
                        $"A table named '{name}' already exists in this workbook. " +
                        "Table names must be unique workbook-wide (case-insensitive).",
                        nameof(name));
                }
            }
        }
    }
}
