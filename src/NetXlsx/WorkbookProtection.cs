// Workbook-protection options record per design §6.2.1 / decision I-54.
// Composed with IWorkbook.Protect to apply structure/windows/revision
// permission flags at the workbook level.

namespace NetXlsx;

/// <summary>
/// Granular workbook-protection options for <see cref="IWorkbook.Protect"/>.
/// <para>
/// Workbook protection guards the workbook's <em>structure</em>
/// (sheet add / delete / rename / reorder / hide), <em>windows</em>
/// (the workbook window's chrome — largely defunct in modern Excel),
/// and <em>revisions</em> (legacy shared-workbook tracking).
/// </para>
/// <para>
/// Like sheet protection (decision I-53), this is a UX guard — not
/// security. Excel's workbook password is hashed with the same
/// known-brute-forceable algorithm.
/// </para>
/// </summary>
public sealed record WorkbookProtection
{
    /// <summary>The empty protection — every flag false.</summary>
    public static WorkbookProtection Default { get; } = new();

    /// <summary>Structure-only protection — the common use case (prevent sheet add/delete/rename).</summary>
    public static WorkbookProtection LockStructure { get; } = new() { Structure = true };

    /// <summary>Lock workbook structure — prevents adding, deleting, renaming, reordering, or hiding sheets.</summary>
    public bool Structure { get; init; }

    /// <summary>Lock workbook windows — Excel 2007 era; largely defunct in modern Excel.</summary>
    public bool Windows { get; init; }

    /// <summary>Lock workbook for revisions — legacy shared-workbook tracking.</summary>
    public bool Revision { get; init; }
}
