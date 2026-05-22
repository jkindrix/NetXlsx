// Centralized reflection over NPOI 2.7.3 internals.
//
// NPOI exposes several capabilities only via protected methods or
// private fields. NetXlsx reaches across them for narrow, well-
// documented operations — concentrated here so:
//
//   1. Each reflection lookup happens once at class-init, not per call.
//   2. The set of "things we reach across NPOI's protection level for"
//      is auditable in one file.
//   3. If an NPOI version bump moves a member, only this file needs
//      to update.
//
// Each accessor names the NPOI symbol it depends on so a future-
// contributor grep against NPOI source finds the link.

using System;
using System.Collections.Generic;
using System.Reflection;
using NPOI;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;

namespace NetXlsx;

internal static class NpoiInternals
{
    // ---- POIXMLDocumentPart.RemoveRelation(POIXMLDocumentPart) -------
    //
    // Protected upstream (NPOI 2.7.3 / .../ooxml/POIXMLDocumentPart.cs).
    // Used by XssfSheet.RemoveTable (decision I-63).

    private static readonly MethodInfo s_removeRelation = typeof(POIXMLDocumentPart).GetMethod(
        name: "RemoveRelation",
        bindingAttr: BindingFlags.NonPublic | BindingFlags.Instance,
        binder: null,
        types: new[] { typeof(POIXMLDocumentPart) },
        modifiers: null)
        ?? throw new InvalidOperationException(
            "NPOI 2.7.3 internal change: POIXMLDocumentPart.RemoveRelation(POIXMLDocumentPart) " +
            "is no longer accessible via reflection. The signature may have changed; update " +
            "src/NetXlsx/Internal/NpoiInternals.cs.");

    internal static void RemoveRelation(POIXMLDocumentPart owner, POIXMLDocumentPart relation)
    {
        s_removeRelation.Invoke(owner, new object?[] { relation });
    }

    // ---- XSSFSheet.tables (Dictionary<string, XSSFTable>) ------------
    //
    // Internal field upstream (NPOI 2.7.3 / .../ooxml/XSSF/UserModel/XSSFSheet.cs).
    // The dictionary is keyed by package-relationship id ("rId..."). NPOI's
    // public GetTables() materializes a list from this dictionary, so a
    // stale entry after RemoveRelation would surface as a phantom in
    // subsequent GetTables() calls.

    private static readonly FieldInfo s_sheetTablesField = typeof(XSSFSheet).GetField(
        name: "tables",
        bindingAttr: BindingFlags.Instance | BindingFlags.NonPublic)
        ?? throw new InvalidOperationException(
            "NPOI 2.7.3 internal change: XSSFSheet.tables field is no longer accessible " +
            "via reflection. Update src/NetXlsx/Internal/NpoiInternals.cs.");

    internal static void RemoveTableFromSheetCache(XSSFSheet sheet, string relationshipId)
    {
        if (s_sheetTablesField.GetValue(sheet) is Dictionary<string, XSSFTable> dict)
        {
            dict.Remove(relationshipId);
        }
    }

    // ---- StylesTable.PutCellStyleXf(CT_Xf) ---------------------------
    // ---- StylesTable.GetCellStyleXfAt(int) ---------------------------
    //
    // Both internal upstream. Used by XssfWorkbook named-style integration
    // (decision I-67). PutCellStyleXf appends to NPOI's private styleXfs
    // list and returns the new size (i.e., new-element-index + 1).
    // GetCellStyleXfAt reads back from that list — needed by the read
    // path during Workbook.Open.
    //
    // Direct manipulation of CT_Stylesheet.cellStyleXfs.xf is not viable
    // because NPOI overwrites that list on save with its internal
    // styleXfs (see StylesTable.cs:887 — `ctSXfs.xf = styleXfs`).

    private static readonly MethodInfo s_putCellStyleXf = typeof(StylesTable).GetMethod(
        name: "PutCellStyleXf",
        bindingAttr: BindingFlags.NonPublic | BindingFlags.Instance,
        binder: null,
        types: new[] { typeof(CT_Xf) },
        modifiers: null)
        ?? throw new InvalidOperationException(
            "NPOI 2.7.3 internal change: StylesTable.PutCellStyleXf(CT_Xf) is no longer " +
            "accessible via reflection. Update src/NetXlsx/Internal/NpoiInternals.cs.");

    private static readonly MethodInfo s_getCellStyleXfAt = typeof(StylesTable).GetMethod(
        name: "GetCellStyleXfAt",
        bindingAttr: BindingFlags.NonPublic | BindingFlags.Instance,
        binder: null,
        types: new[] { typeof(int) },
        modifiers: null)
        ?? throw new InvalidOperationException(
            "NPOI 2.7.3 internal change: StylesTable.GetCellStyleXfAt(int) is no longer " +
            "accessible via reflection. Update src/NetXlsx/Internal/NpoiInternals.cs.");

    /// <summary>
    /// Appends <paramref name="xf"/> to NPOI's internal cellStyleXfs
    /// list. Returns the new size (so <c>result - 1</c> is the index
    /// of the just-added entry).
    /// </summary>
    internal static int PutCellStyleXf(StylesTable styles, CT_Xf xf)
    {
        var result = s_putCellStyleXf.Invoke(styles, new object[] { xf });
        return (int)(result ?? 0);
    }

    /// <summary>
    /// Returns the cellStyleXfs entry at <paramref name="index"/>, or
    /// <c>null</c> if the index is out of range.
    /// </summary>
    internal static CT_Xf? GetCellStyleXfAt(StylesTable styles, int index)
    {
        return s_getCellStyleXfAt.Invoke(styles, new object[] { index }) as CT_Xf;
    }

    // ---- StylesTable.PutCellXf(CT_Xf) -------------------------------
    //
    // Internal upstream. Adds a CT_Xf to the regular cellXfs table.
    // Used at named-style read time to materialize the cellStyleXfs
    // entry as a regular cell-style so the existing CellStylePool.
    // ReadFromNpoi parser can run against it.

    private static readonly MethodInfo s_putCellXf = typeof(StylesTable).GetMethod(
        name: "PutCellXf",
        bindingAttr: BindingFlags.NonPublic | BindingFlags.Instance,
        binder: null,
        types: new[] { typeof(CT_Xf) },
        modifiers: null)
        ?? throw new InvalidOperationException(
            "NPOI 2.7.3 internal change: StylesTable.PutCellXf(CT_Xf) is no longer " +
            "accessible via reflection. Update src/NetXlsx/Internal/NpoiInternals.cs.");

    internal static int PutCellXf(StylesTable styles, CT_Xf xf)
    {
        var result = s_putCellXf.Invoke(styles, new object[] { xf });
        return (int)(result ?? 0);
    }
}
