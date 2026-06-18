namespace NetXlsx;

/// <summary>
/// What kind of sheet an <see cref="ISheet"/> represents (decision I-92).
/// Most sheets are <see cref="Worksheet"/> (a cell grid). A workbook opened
/// from a file may also contain a <see cref="Chartsheet"/> (a full-window
/// chart with no grid) or, rarely, a legacy <see cref="Dialogsheet"/>; those
/// open as placeholders — they appear in the sheet collection and round-trip,
/// but grid access throws (see <see cref="ISheet.Kind"/>).
/// </summary>
public enum SheetKind
{
    /// <summary>A normal worksheet with a cell grid. Created by <see cref="IWorkbook.AddSheet(string)"/>.</summary>
    Worksheet = 0,

    /// <summary>A chartsheet — a full-window chart with no cell grid. Open/round-trip only; grid access throws.</summary>
    Chartsheet = 1,

    /// <summary>A legacy Excel-5.0/95 dialog sheet. Open/round-trip only; grid access throws.</summary>
    Dialogsheet = 2,
}
