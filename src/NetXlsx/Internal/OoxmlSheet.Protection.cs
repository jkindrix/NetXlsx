// I-82 engine swap — Open XML SDK-backed sheet protection (structure slice, 5b).
//
// Mirrors the NPOI engine's ISheet.Protect / Unprotect / IsProtected contract
// (decision I-53) on the SDK engine, writing the OOXML <sheetProtection> node
// directly. <sheetProtection> sits between <sheetData> and <mergeCells> in
// CT_Worksheet, so it is placed by OoxmlSchemaOrder (which orders it correctly
// even when a merge already exists; SDK-quirk #3 / #8).
//
// Like the NPOI engine, Protect sets @sheet=true plus all 15 granular lock flags
// explicitly from the SheetProtection options, so the result never depends on the
// schema's per-attribute defaults. An optional password is stored as the legacy
// 16-bit XOR verifier in @password (ST_UnsignedShortHex) — the same weak hash
// Excel and NPOI write; a UX guard, not security (documented on SheetProtection).

using System;
using System.Globalization;
using DocumentFormat.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    public void Protect(string? password = null, SheetProtection? options = null)
    {
        _workbook.ThrowIfDisposed();
        var opts = options ?? SheetProtection.Default;

        var sp = GetOrCreateSheetProtection();
        sp.Sheet = true;
        sp.Password = password is null ? null : HexBinaryValue.FromString(LegacyPasswordHash(password));

        // Every flag set explicitly from opts (NPOI's net behavior): @sheet stays
        // true; objects/scenarios reflect the options like every other lock.
        sp.FormatCells = opts.LockFormatCells;
        sp.FormatColumns = opts.LockFormatColumns;
        sp.FormatRows = opts.LockFormatRows;
        sp.InsertColumns = opts.LockInsertColumns;
        sp.InsertRows = opts.LockInsertRows;
        sp.InsertHyperlinks = opts.LockInsertHyperlinks;
        sp.DeleteColumns = opts.LockDeleteColumns;
        sp.DeleteRows = opts.LockDeleteRows;
        sp.SelectLockedCells = opts.LockSelectLockedCells;
        sp.SelectUnlockedCells = opts.LockSelectUnlockedCells;
        sp.Sort = opts.LockSort;
        sp.AutoFilter = opts.LockAutoFilter;
        sp.PivotTables = opts.LockPivotTables;
        sp.Objects = opts.LockObjects;
        sp.Scenarios = opts.LockScenarios;
    }

    public void Unprotect()
    {
        _workbook.ThrowIfDisposed();
        // Mirrors NPOI's ProtectSheet(null): remove the element entirely rather
        // than leave a @sheet=false stub.
        Worksheet.GetFirstChild<S.SheetProtection>()?.Remove();
    }

    public bool IsProtected
    {
        get
        {
            _workbook.ThrowIfDisposed();
            return Worksheet.GetFirstChild<S.SheetProtection>()?.Sheet?.Value ?? false;
        }
    }

    private S.SheetProtection GetOrCreateSheetProtection()
        => OoxmlSchemaOrder.GetOrInsert(Worksheet, static () => new S.SheetProtection());

    // The standard legacy 16-bit Excel password verifier (the classic hash Excel
    // writes into @password for sheet protection). Returned as the 4-hex-digit
    // ST_UnsignedShortHex form the attribute expects. Excel's sheet-protection
    // password is brute-forceable by design — this is a UX guard only, and the
    // contract asserts only presence + round-trip + schema-validity, not a
    // byte-for-byte match with NPOI's verifier.
    internal static string LegacyPasswordHash(string password)
    {
        int hash = 0;
        if (password.Length > 0)
        {
            for (int i = password.Length - 1; i >= 0; i--)
            {
                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
                hash ^= password[i];
            }
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
            hash ^= password.Length;
            hash ^= 0xCE4B;
        }
        return hash.ToString("X4", CultureInfo.InvariantCulture);
    }
}
