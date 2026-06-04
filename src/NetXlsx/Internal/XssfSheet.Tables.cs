// XssfSheet — Excel Tables (ListObject) surface per decision I-51.
// Core class structure is in XssfSheet.cs.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using NPOI;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal sealed partial class XssfSheet
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

        // Collect header cell values from the first row of the range —
        // these become the table's column names. Excel requires a non-empty
        // string in each header cell; reject blanks loud.
        var headers = new List<string>(c2 - c1 + 1);
        for (int col = c1; col <= c2; col++)
        {
            var cell = (XSSFCell?)_underlying
                .GetRow(r1 - 1)?
                .GetCell(col - 1);
            var text = cell?.CellType == CellType.String ? cell.StringCellValue : null;
            if (string.IsNullOrEmpty(text))
            {
                throw new ArgumentException(
                    $"Header cell at {CellAddress.Format(r1, col)} is empty or non-string. " +
                    "Set header values on every column before calling AddTable.",
                    nameof(a1Range));
            }
            // Excel requires column names to be unique within a table
            // (case-insensitive). NPOI enforces this in CreateColumn but
            // its error is awkward; pre-check for a friendlier message.
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

        // Build the NPOI table. XSSFTable.CreateColumn in NPOI 2.7.3
        // throws ArgumentOutOfRangeException because the underlying
        // CT_TableColumns.tableColumn list is uninitialized; bypass it
        // by populating the CT_Table directly. Captured in
        // implementation-notes.md (NPOI 2.7.3 surprise).
        var npoi = _underlying.CreateTable();
        npoi.Name = name;
        npoi.DisplayName = name;

        var ctTable = npoi.GetCTTable();
        var ctColumns = new CT_TableColumns
        {
            tableColumn = new List<CT_TableColumn>(headers.Count),
            count = (uint)headers.Count,
        };
        ctTable.tableColumns = ctColumns;
        for (int i = 0; i < headers.Count; i++)
        {
            var ctCol = ctColumns.InsertNewTableColumn(i);
            ctCol.id = (uint)(i + 1);
            ctCol.name = headers[i];
        }

        npoi.SetCellReferences(new AreaReference(
            new CellReference(r1 - 1, c1 - 1),
            new CellReference(r2 - 1, c2 - 1),
            SpreadsheetVersion.EXCEL2007));

        if (style is not null)
        {
            npoi.StyleName = style;
        }

        return new XssfTable(_workbook, this, npoi);
    }

    public IReadOnlyList<ITable> Tables
    {
        get
        {
            _workbook.ThrowIfDisposed();
            var list = _underlying.GetTables();
            if (list is null || list.Count == 0) return Array.Empty<ITable>();
            var result = new ITable[list.Count];
            for (int i = 0; i < list.Count; i++)
                result[i] = new XssfTable(_workbook, this, list[i]);
            return result;
        }
    }

    public bool TryGetTable(string name, [MaybeNullWhen(false)] out ITable table)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        var list = _underlying.GetTables();
        if (list is not null)
        {
            foreach (var t in list)
            {
                if (string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    table = new XssfTable(_workbook, this, t);
                    return true;
                }
            }
        }
        table = null;
        return false;
    }

    public void RemoveTable(ITable table)
    {
        _workbook.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(table);
        if (table is not XssfTable xt)
        {
            throw new ArgumentException(
                "Table instance is not an XssfTable — likely a foreign or mocked implementation.",
                nameof(table));
        }
        // Verify the table actually belongs to this sheet. NPOI's
        // GetRelationId returns null if the part isn't a relation of
        // this sheet — that's the cheapest belongs-to check available.
        var relId = _underlying.GetRelationId(xt.Npoi);
        if (string.IsNullOrEmpty(relId))
        {
            throw new ArgumentException(
                "Table does not belong to this sheet (no relation found).",
                nameof(table));
        }

        // Step 1: drop the <tablePart> entry referencing this relId
        // from CT_Worksheet.tableParts. Walk the list backwards in
        // case duplicate ids ever appear (defensive — shouldn't, but
        // the cost is one extra comparison).
        var ws = _underlying.GetCTWorksheet();
        if (ws.tableParts?.tablePart is { } parts)
        {
            for (int i = parts.Count - 1; i >= 0; i--)
            {
                if (parts[i].id == relId) parts.RemoveAt(i);
            }
            ws.tableParts.count = (uint)parts.Count;
            // OOXML allows the <tableParts> element to be present-but-
            // empty, but Excel is friendlier with it absent entirely
            // when the sheet has no remaining tables.
            if (parts.Count == 0) ws.tableParts = null;
        }

        // Step 2: drop the package relationship + part via
        // POIXMLDocumentPart.RemoveRelation. The method is protected
        // upstream; reach across via the cached MethodInfo from
        // NpoiInternals (one reflection lookup, shared across calls).
        NpoiInternals.RemoveRelation(_underlying, xt.Npoi);

        // Step 3: drop the cached entry from XSSFSheet's internal
        // `tables` dictionary so a subsequent GetTables() snapshot
        // doesn't surface the removed table.
        NpoiInternals.RemoveTableFromSheetCache(_underlying, relId);
    }

    // ---- Internal helpers ---------------------------------------------

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
        // Reject anything that would parse as a column-letter+row-number
        // cell reference (e.g. "A1", "ABC123") — Excel rejects these names.
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
        // Per Excel: table names share the namespace with named ranges
        // and with other tables in the workbook. Check both.
        foreach (var nr in _workbook.Npoi.GetAllNames())
        {
            if (string.Equals(nr.NameName, name, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(
                    $"Name '{name}' is already taken by a named range " +
                    "(table and named-range names share the workbook-wide namespace).",
                    nameof(name));
            }
        }
        for (int s = 0; s < _workbook.Npoi.NumberOfSheets; s++)
        {
            var sheet = _workbook.Npoi.GetSheetAt(s) as XSSFSheet;
            if (sheet is null) continue;
            var tables = sheet.GetTables();
            if (tables is null) continue;
            foreach (var t in tables)
            {
                if (string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    throw new ArgumentException(
                        $"A table named '{name}' already exists on sheet '{sheet.SheetName}'. " +
                        "Table names must be unique workbook-wide (case-insensitive).",
                        nameof(name));
                }
            }
        }
    }
}
