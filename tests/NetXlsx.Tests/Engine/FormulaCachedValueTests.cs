// R-7 / design I7 conformance: formula-cached results are readable across
// the entire typed-getter surface, and GetString returns the error literal
// for error cells. Pre-fix, KindOf short-circuited every formula cell to
// CellKind.Formula before checking `t`, and no getter had a Formula/Error
// arm — GetNumber on <f>1+1</f><v>2</v> was null, GetString on a cached
// formula-string was "", and the generated ReadRows threw on any
// Excel-authored file with a formula column.
//
// This engine's own writer never stores a cached <v> (design #46), so the
// cells are crafted through the Underlying escape hatch — the same shapes
// Excel/LibreOffice author — and the load-bearing cases re-verify through
// a full save → reopen so the from-file read path is what's pinned.

using System;
using System.Collections;
using System.IO;
using System.Linq;
using AwesomeAssertions;
using NetXlsx;
using NetXlsx.Tests.SourceGen;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace NetXlsx.Tests.Engine;

public class FormulaCachedValueTests
{
    // Authors <c><f>…</f><v>cached</v></c> (+ optional t) — the shape a
    // calculating producer writes. SetFormula deliberately emits no <v>.
    private static void SetCachedFormula(ICell cell, string formula, string cached, S.CellValues? t = null)
    {
        cell.SetFormula(formula);
        var raw = cell.Underlying;
        if (t is { } dataType) raw.DataType = dataType;
        raw.CellValue = new S.CellValue(cached);
    }

    private static IWorkbook Reopen(IWorkbook wb)
    {
        var ms = new MemoryStream();
        wb.Save(ms);
        wb.Dispose();
        ms.Position = 0;
        return Workbook.Open(ms);
    }

    [Fact]
    public void GetNumber_Reads_Cached_Numeric_Formula_Result()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        SetCachedFormula(sheet["A1"], "=1+1", "2");

        sheet["A1"].Kind.Should().Be(CellKind.Formula, "the public Kind contract is unchanged");
        sheet["A1"].GetNumber().Should().Be(2.0);
        sheet["A1"].GetString().Should().Be("2", "I7: formula → cached result as string");

        using var read = Reopen(wb);
        read["S"]["A1"].GetNumber().Should().Be(2.0, "the from-file read path must agree");
        read["S"]["A1"].GetFormula().Should().Be("=1+1");
    }

    [Fact]
    public void GetString_Reads_Cached_String_Formula_Result()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        // t="str" — the cached-string shape Excel writes for =CONCAT etc.
        SetCachedFormula(sheet["A1"], "=CONCAT(\"he\",\"llo\")", "hello", S.CellValues.String);

        sheet["A1"].GetString().Should().Be("hello");
        sheet["A1"].GetNumber().Should().BeNull("a cached string result is not a number");

        using var read = Reopen(wb);
        read["S"]["A1"].GetString().Should().Be("hello");
    }

    [Fact]
    public void GetBool_Reads_Cached_Boolean_Formula_Result()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        // LibreOffice authors t="b" + <f/> + <v>1</v> for boolean formulas.
        SetCachedFormula(sheet["A1"], "=1=1", "1", S.CellValues.Boolean);

        sheet["A1"].GetBool().Should().BeTrue();
        sheet["A1"].GetNumber().Should().Be(1.0, "bool reads as 1/0 through GetNumber, same as plain bool cells");
        sheet["A1"].GetString().Should().Be("TRUE");

        using var read = Reopen(wb);
        read["S"]["A1"].GetBool().Should().BeTrue();
    }

    [Fact]
    public void GetString_Returns_Error_Literal_For_Error_Cells()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        SetCachedFormula(sheet["A1"], "=1/0", "#DIV/0!", S.CellValues.Error);

        sheet["A1"].GetString().Should().Be("#DIV/0!", "I7 verbatim: error → error code");
        sheet["A1"].GetError().Should().Be(CellError.DivByZero, "GetError (#49) is unchanged");
        sheet["A1"].GetNumber().Should().BeNull();
        sheet["A1"].GetBool().Should().BeNull();

        using var read = Reopen(wb);
        read["S"]["A1"].GetString().Should().Be("#DIV/0!");
    }

    [Fact]
    public void GetDate_And_GetTime_Read_Cached_Serial_On_DateFormatted_Formula_Cell()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        // 2026-06-03 09:30:15 as a 1900-system serial, cached on a formula
        // cell that carries a date format — Excel's =NOW() shape.
        var instant = new DateTime(2026, 6, 3, 9, 30, 15);
        SetCachedFormula(sheet["A1"], "=NOW()", "46176.396006944444444");
        sheet["A1"].NumberFormat("yyyy-mm-dd hh:mm:ss");

        sheet["A1"].GetDate().Should().Be(instant);

        // A time-of-day formula (serial in [0, 1)) reads through GetTime —
        // and picks up the R-6 millisecond rounding on the 15-digit cache.
        SetCachedFormula(sheet["A2"], "=TIME(9,30,15)", "0.396006944444444");
        sheet["A2"].GetTime().Should().Be(new TimeOnly(9, 30, 15), "GetTime flows through GetNumber");
    }

    [Fact]
    public void Fresh_Formula_Without_Cached_Value_Still_Reads_As_Nothing()
    {
        // This engine's own SetFormula stores no <v> (design #46) — the
        // pre-R-7 behavior for THAT shape was correct and must not change.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetFormula("=1+1");

        sheet["A1"].Kind.Should().Be(CellKind.Formula);
        sheet["A1"].GetNumber().Should().BeNull();
        sheet["A1"].GetString().Should().BeEmpty();
        sheet["A1"].GetBool().Should().BeNull();
        sheet["A1"].GetDate().Should().BeNull();
    }

    [Fact]
    public void Empty_Cached_V_Element_Reads_As_Nothing()
    {
        // NPOI authors an empty <v/> under fresh formulas — both forms mean
        // "no cached value" and must read identically.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        SetCachedFormula(sheet["A1"], "=1+1", "");

        sheet["A1"].GetNumber().Should().BeNull();
        sheet["A1"].GetString().Should().BeEmpty();
    }

    [Fact]
    public void Generated_ReadRows_Imports_A_Formula_Column()
    {
        // The R-7 knock-on: emitted ReadRows does `GetNumber() ?? throw`,
        // so a typed import over any calculating producer's file threw.
        // Compile-and-run proof over a workbook with a cached formula column.
        const string src = @"
using NetXlsx;
namespace T;
[Worksheet]
public partial record LineItem(
    [property: Column(""Name"")]  string Name,
    [property: Column(""Total"")] double Total);";
        var output = GeneratorHarness.RunWithFullReferences(src, out var compilation);
        output.CompilationDiagnostics.Should().NotContain(
            d => d.Severity == Microsoft.CodeAnalysis.DiagnosticSeverity.Error);
        var asm = GeneratorHarness.EmitAndLoad(compilation);
        var ext = asm.GetType("T.LineItem_SheetExtensions")!;

        using var ms = new MemoryStream();
        using (var wb = Workbook.Create())
        {
            var sheet = wb.AddSheet("Items");
            sheet.AppendRow().Set(1, "Name").Set(2, "Total");
            sheet.AppendRow().Set(1, "widgets");
            SetCachedFormula(sheet["B2"], "=2*21", "42");
            sheet.AppendRow().Set(1, "gadgets");
            SetCachedFormula(sheet["B3"], "=SUM(1,2)", "3");
            wb.Save(ms);
        }

        ms.Position = 0;
        using (var wb = Workbook.Open(ms))
        {
            var rows = (IEnumerable)ext.GetMethod("ReadRows")!.Invoke(null, new object?[] { wb["Items"], 1 })!;
            var back = rows.Cast<object>().ToList();

            back.Should().HaveCount(2);
            var totalProp = back[0].GetType().GetProperty("Total")!;
            totalProp.GetValue(back[0]).Should().Be(42.0);
            totalProp.GetValue(back[1]).Should().Be(3.0);
        }
    }
}
