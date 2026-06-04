// Engine-agnostic OOXML inspection for behavioral tests (I-82 cutover
// phase 1). Where the public API has no read-back for a surface (granular
// protection flags, pane state, outline levels, autofilter shapes, …),
// tests assert on the persisted file format itself: save the workbook to
// memory, open the zip, parse the part XML. The OOXML bytes are the real
// contract, so these assertions hold identically on the NPOI and Open XML
// SDK engines — unlike `.Underlying` reach-throughs, which are
// engine-typed and break at the v2.0.0 cutover.
//
// Attribute conventions: OOXML booleans serialize as "1"/"true" (engines
// differ); absent attributes mean the schema default (false for the
// surfaces asserted here). Use BoolAttr for both.

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace NetXlsx.Tests;

internal static class SavedOoxml
{
    /// <summary>The main spreadsheetml namespace (x:).</summary>
    public static readonly XNamespace Main =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>The spreadsheet drawing namespace (xdr:).</summary>
    public static readonly XNamespace Xdr =
        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    /// <summary>The drawingml main namespace (a:).</summary>
    public static readonly XNamespace Dml =
        "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Saves <paramref name="wb"/> to memory and returns the XML of the
    /// given OPC part (e.g. <c>xl/workbook.xml</c>).
    /// </summary>
    public static XDocument Part(IWorkbook wb, string partPath)
    {
        using var ms = new MemoryStream();
        wb.Save(ms);
        ms.Position = 0;
        return ReadPart(ms, partPath);
    }

    /// <summary>Returns the XML of an OPC part from an .xlsx file on disk.</summary>
    public static XDocument PartFromFile(string path, string partPath)
    {
        using var fs = File.OpenRead(path);
        return ReadPart(fs, partPath);
    }

    /// <summary><c>xl/workbook.xml</c> of the saved workbook.</summary>
    public static XDocument WorkbookXml(IWorkbook wb) => Part(wb, "xl/workbook.xml");

    /// <summary><c>xl/styles.xml</c> of the saved workbook.</summary>
    public static XDocument StylesXml(IWorkbook wb) => Part(wb, "xl/styles.xml");

    /// <summary>
    /// <c>xl/worksheets/sheetN.xml</c> of the saved workbook (1-based,
    /// in sheet creation order).
    /// </summary>
    public static XDocument SheetXml(IWorkbook wb, int sheetNumber = 1)
        => Part(wb, $"xl/worksheets/sheet{sheetNumber}.xml");

    /// <summary><c>xl/drawings/drawingN.xml</c> of the saved workbook.</summary>
    public static XDocument DrawingXml(IWorkbook wb, int drawingNumber = 1)
        => Part(wb, $"xl/drawings/drawing{drawingNumber}.xml");

    /// <summary>
    /// Reads an OOXML boolean attribute: <c>"1"</c> or <c>"true"</c> is
    /// true; anything else — including an absent attribute or a null
    /// element — is false (the schema default for these surfaces).
    /// </summary>
    public static bool BoolAttr(XElement? element, string attribute)
        => element?.Attribute(attribute)?.Value is "1" or "true";

    /// <summary>
    /// The <c>&lt;c&gt;</c> element for an A1 address in a saved sheet
    /// XML, or null when the cell was never materialized.
    /// </summary>
    public static XElement? Cell(XDocument sheetXml, string a1Address)
        => sheetXml.Root!.Element(Main + "sheetData")!
            .Elements(Main + "row")
            .SelectMany(r => r.Elements(Main + "c"))
            .FirstOrDefault(c => (string?)c.Attribute("r") == a1Address);

    /// <summary>
    /// The cell's persisted style index (<c>c/@s</c> into cellXfs);
    /// null when absent (= xf 0, the default).
    /// </summary>
    public static int? CellStyleIndex(XDocument sheetXml, string a1Address)
        => (int?)Cell(sheetXml, a1Address)?.Attribute("s");

    private static XDocument ReadPart(Stream zipStream, string partPath)
    {
        using var zip = new ZipArchive(zipStream, ZipArchiveMode.Read, leaveOpen: true);
        var entry = zip.GetEntry(partPath)
            ?? throw new InvalidOperationException(
                $"Part '{partPath}' not found in the saved workbook.");
        using var s = entry.Open();
        return XDocument.Load(s);
    }
}
