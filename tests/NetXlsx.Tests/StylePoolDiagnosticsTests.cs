// Coverage for the post-v1.1 style-pool diagnostics surface
// (decision I-61): IWorkbook.GetStylePoolDiagnostics returns
// non-allocating counter snapshots; hit/miss reflect actual pool
// activity; dedup ratios fall in [0, 1].

using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class StylePoolDiagnosticsTests
{
    [Fact]
    public void Fresh_Workbook_Reports_Zero_Counts()
    {
        using var wb = Workbook.Create();
        var d = wb.GetStylePoolDiagnostics();
        d.StyleHitCount.Should().Be(0);
        d.StyleMissCount.Should().Be(0);
        d.FontHitCount.Should().Be(0);
        d.FontMissCount.Should().Be(0);
        d.UniqueStyles.Should().Be(0);
        d.UniqueFonts.Should().Be(0);
    }

    [Fact]
    public void Distinct_Styles_Increment_Miss_Counter()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].Style(new CellStyle { Bold = true });
        sh["A2"].Style(new CellStyle { Italic = true });
        sh["A3"].Style(new CellStyle { Underline = UnderlineStyle.Single });

        var d = wb.GetStylePoolDiagnostics();
        d.StyleMissCount.Should().Be(3);
        d.UniqueStyles.Should().Be(3);
    }

    [Fact]
    public void Repeated_Style_Reuses_Pool_Entry()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        var style = new CellStyle { Bold = true, FontSize = 14 };
        sh["A1"].Style(style);
        sh["A2"].Style(style);
        sh["A3"].Style(style);

        var d = wb.GetStylePoolDiagnostics();
        d.UniqueStyles.Should().Be(1, "one distinct CellStyle => one pool entry");
        d.StyleHitCount.Should().Be(2, "second and third Style() calls hit");
        d.StyleMissCount.Should().Be(1, "first Style() call is a miss");
        d.StyleDedupRatio.Should().BeApproximately(2.0 / 3, 0.001);
    }

    [Fact]
    public void DedupRatio_Returns_Zero_When_No_Lookups()
    {
        using var wb = Workbook.Create();
        var d = wb.GetStylePoolDiagnostics();
        d.StyleDedupRatio.Should().Be(0);
        d.FontDedupRatio.Should().Be(0);
    }

    [Fact]
    public void Diagnostics_Is_A_Snapshot_Not_A_Live_View()
    {
        using var wb = Workbook.Create();
        var sh = wb.AddSheet("S");
        sh["A1"].Style(new CellStyle { Bold = true });

        var snap = wb.GetStylePoolDiagnostics();
        var initialUnique = snap.UniqueStyles;

        sh["A2"].Style(new CellStyle { Italic = true });

        snap.UniqueStyles.Should().Be(initialUnique,
            "the captured snapshot must not change when the underlying pool grows");
        wb.GetStylePoolDiagnostics().UniqueStyles.Should().Be(initialUnique + 1,
            "but a fresh snapshot reflects the new pool size");
    }
}
