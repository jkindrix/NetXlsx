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
}
