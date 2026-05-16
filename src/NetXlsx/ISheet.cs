// Public marker interface for the source generator. The full public surface
// (per docs/design.md §6.4) lands in the next implementation milestone.
// This stub exists so source-generated extension methods on ISheet can
// compile against v0.1.x while the rest of the API is being built out.

namespace NetXlsx;

/// <summary>
/// Represents a worksheet within an <see cref="IWorkbook"/>. The full
/// surface — cell access, row enumeration, range operations, merging,
/// freeze panes, etc. — is defined in <c>docs/design.md §6.4</c> and
/// lands in the next implementation milestone. This interface is
/// deliberately empty at v0.1.x: it exists so that source-generated
/// typed-mapping extension methods can name <c>ISheet</c> as their
/// extension target while the body wiring is being built.
/// </summary>
public interface ISheet
{
    // Empty by design. Real members specified in design §6.4 land in the
    // milestone that builds out IWorkbook / IRow / ICell. The source
    // generator's emitted method bodies throw NotImplementedException
    // until that milestone, by which point this interface gains the
    // members the bodies need.
}

/// <summary>
/// Represents an Excel workbook. Empty by design at v0.1.x for the same
/// reason as <see cref="ISheet"/>: the full surface (per <c>docs/design.md
/// §6.2</c>) lands in the next implementation milestone. Source-generated
/// extension methods reference this type as a marker so they can compile
/// against v0.1.x without depending on members that don't exist yet.
/// </summary>
public interface IWorkbook
{
    // Empty by design — see XML doc above.
}
