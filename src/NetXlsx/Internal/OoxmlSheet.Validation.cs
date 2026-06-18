// I-82 engine swap — Open XML SDK-backed data validation (CF/validation/
// tables/autofilter/sort slice). Surface per decision I-55.
//
// OOXML stores validations as a single <dataValidations count="N"> worksheet
// child (0..1) holding one <dataValidation> per rule: @type (list/whole/
// decimal/date/textLength/custom), @operator, @allowBlank, @sqref, and
// <formula1>/<formula2> children. <dataValidations> sits in CT_Worksheet's
// strict child sequence (after <conditionalFormatting>, before <hyperlinks>),
// so the insert routes through OoxmlSchemaOrder (SDK-quirk #8); it is a 0..1
// singleton, so GetOrInsert applies directly. The API only ADDS rules (no
// remove), so the empty-container concern (SDK-quirk #7) cannot arise.
//
// The rule axes come from DataValidation's engine-agnostic descriptor
// (Kind/Operator/Formula1/Formula2) — the NPOI-typed Build closure is never
// touched here. Oracle-checked against the NPOI engine's emitted XML
// (SDK-quirk #11 habit): allowBlank="1" is always set (it is NOT the schema
// default); NPOI's errorStyle="stop" / imeMode="noControl" /
// showDropDown="0" / showInputMessage="0" / showErrorMessage="0" /
// operator="between" on list+custom are all schema DEFAULTS and are omitted;
// @count is kept in sync (Excel and NPOI both emit it).
//
// RemoveValidation (I-91 removal family) is key-based: it matches the rule by
// its exact (canonical) range and drops the container + fixes @count when the
// last rule goes (SDK-quirk #7). It checks BOTH the plain <dataValidations>
// container NetXlsx authors and the x14 <extLst> form opened files carry for
// cross-sheet list sources ([A-2026-06-11] dual storage).

using System;
using System.Linq;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace NetXlsx;

internal sealed partial class OoxmlSheet
{
    public void AddValidation(string a1Range, DataValidation validation)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(a1Range);
        ArgumentNullException.ThrowIfNull(validation);

        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);

        var dv = new S.DataValidation
        {
            Type = ToValidationType(validation.Kind),
            AllowBlank = true,
            // 1x1 collapses to "A2" (NPOI's CellRangeAddressList parity).
            SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue>
            {
                InnerText = CellAddress.FormatRange(r1, c1, r2, c2),
            },
        };
        if (validation.Operator != DataValidation.CompareOp.None)
            dv.Operator = ToValidationOperator(validation.Operator);

        dv.AppendChild(new S.Formula1(validation.Formula1));
        if (validation.Formula2 is not null)
            dv.AppendChild(new S.Formula2(validation.Formula2));

        var container = OoxmlSchemaOrder.GetOrInsert(Worksheet, static () => new S.DataValidations());
        container.AppendChild(dv);
        container.Count = (uint)container.Elements<S.DataValidation>().Count();
    }

    public void RemoveValidation(string a1Range)
    {
        ThrowIfUnusable();
        ArgumentNullException.ThrowIfNull(a1Range);

        var (r1, c1, r2, c2) = CellAddress.ParseRange(a1Range);
        var canonical = CellAddress.FormatRange(r1, c1, r2, c2);

        if (TryRemovePlainValidation(canonical)) return;
        if (TryRemoveX14Validation(canonical)) return;

        // Key-based, RemoveTable/RemoveNamedRange semantics: an absent key
        // throws (deliberately NOT the silent no-op of the grandfathered
        // UnmergeCells; the cell-level removals are the idempotent ones).
        throw new ArgumentException(
            $"no data validation has the exact range '{canonical}' on this sheet.", nameof(a1Range));
    }

    // The plain <dataValidations> form (what NetXlsx authors). The range key is
    // the <dataValidation @sqref> value.
    private bool TryRemovePlainValidation(string canonical)
    {
        var container = Worksheet.GetFirstChild<S.DataValidations>();
        if (container is null) return false;
        var match = container.Elements<S.DataValidation>()
            .FirstOrDefault(dv => SqrefMatches(dv.SequenceOfReferences?.InnerText, canonical));
        if (match is null) return false;

        match.Remove();
        int remaining = container.Elements<S.DataValidation>().Count();
        if (remaining == 0) container.Remove();
        else container.Count = (uint)remaining;
        return true;
    }

    // The x14 form ([A-2026-06-11] dual storage): cross-sheet list-source rules
    // live in <x14:dataValidations> inside the worksheet <extLst>, keyed by an
    // <xm:sqref> child rather than a @sqref attribute. An emptied x14 container
    // is removed together with its <ext> wrapper — an emptied-but-present ext
    // also triggers Excel repair (ClosedXML #2594).
    private bool TryRemoveX14Validation(string canonical)
    {
        foreach (var container in Worksheet.Descendants<X14.DataValidations>().ToList())
        {
            var match = container.Elements<X14.DataValidation>()
                .FirstOrDefault(dv => SqrefMatches(X14SqrefOf(dv), canonical));
            if (match is null) continue;

            match.Remove();
            int remaining = container.Elements<X14.DataValidation>().Count();
            if (remaining > 0)
            {
                container.Count = (uint)remaining;
                return true;
            }

            // The wrapping <ext>/<extLst> are WorksheetExtension/
            // WorksheetExtensionList under a worksheet (distinct SDK types from
            // the generic Extension/ExtensionList) — match by local name so the
            // exact type is irrelevant.
            var ext = container.Ancestors().FirstOrDefault(e => e.LocalName == "ext");
            container.Remove();
            if (ext is not null)
            {
                var extLst = ext.Parent;
                ext.Remove();
                if (extLst is not null && extLst.ChildElements.Count == 0) extLst.Remove();
            }
            return true;
        }
        return false;
    }

    // The x14 rule's range key is its <xm:sqref> child; read by local name so
    // the exact CLR type of the element is irrelevant.
    private static string? X14SqrefOf(X14.DataValidation dv)
        => dv.ChildElements.FirstOrDefault(e => e.LocalName == "sqref")?.InnerText;

    // Exact-range match by canonical form: a direct string hit, or a
    // single-range sqref that canonicalizes to the same range (so "A1" matches
    // a stored "A1:A1" and vice versa). Space-separated multi-range sqrefs only
    // match verbatim.
    private static bool SqrefMatches(string? sqref, string canonical)
    {
        if (sqref is null) return false;
        var trimmed = sqref.Trim();
        if (trimmed == canonical) return true;
        if (trimmed.Length == 0 || trimmed.Contains(' ')) return false;
        try
        {
            var (r1, c1, r2, c2) = CellAddress.ParseRange(trimmed);
            return CellAddress.FormatRange(r1, c1, r2, c2) == canonical;
        }
        catch (InvalidCellAddressException)
        {
            return false;
        }
    }

    private static S.DataValidationValues ToValidationType(DataValidation.ConstraintKind kind) => kind switch
    {
        DataValidation.ConstraintKind.List => S.DataValidationValues.List,
        DataValidation.ConstraintKind.Whole => S.DataValidationValues.Whole,
        DataValidation.ConstraintKind.Decimal => S.DataValidationValues.Decimal,
        DataValidation.ConstraintKind.Date => S.DataValidationValues.Date,
        DataValidation.ConstraintKind.TextLength => S.DataValidationValues.TextLength,
        _ => S.DataValidationValues.Custom,
    };

    private static S.DataValidationOperatorValues ToValidationOperator(DataValidation.CompareOp op) => op switch
    {
        DataValidation.CompareOp.Between => S.DataValidationOperatorValues.Between,
        DataValidation.CompareOp.Equal => S.DataValidationOperatorValues.Equal,
        DataValidation.CompareOp.GreaterThan => S.DataValidationOperatorValues.GreaterThan,
        DataValidation.CompareOp.LessThan => S.DataValidationOperatorValues.LessThan,
        DataValidation.CompareOp.LessThanOrEqual => S.DataValidationOperatorValues.LessThanOrEqual,
        _ => S.DataValidationOperatorValues.GreaterThanOrEqual,
    };
}
