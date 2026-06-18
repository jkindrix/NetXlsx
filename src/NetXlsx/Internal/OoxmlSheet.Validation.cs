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

using System;
using System.Linq;
using S = DocumentFormat.OpenXml.Spreadsheet;

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
