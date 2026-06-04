// I-82 engine swap — IRange contract parity (pre-cutover parity slice).
//
// The O-15 scout found two OoxmlRange divergences from the NPOI engine's
// pinned v1 contracts: the bounds properties (FirstRow/FirstCol/LastRow/
// LastCol) lacked disposed-workbook guards (decision #42, pinned by
// DisposedWorkbookMatrixTests), and IRange.Value's unsupported-type message
// diverged from the pinned "is not a supported scalar" wording (and the
// unsigned scalar cases ushort/uint/ulong were missing entirely).

using System;
using AwesomeAssertions;
using NetXlsx;
using Xunit;

namespace NetXlsx.OoxmlEngine.Tests;

public class RangeContractParityTests
{
    // ---- Disposed guards on the bounds (decision #42) -----------------

    [Fact]
    public void Range_Bounds_Throw_ObjectDisposedException_After_Workbook_Dispose()
    {
        var wb = Workbook.CreateOoxml();
        var range = wb.AddSheet("S").Range("B2:D5");
        wb.Dispose();

        ((Action)(() => { _ = range.FirstRow; })).Should().Throw<ObjectDisposedException>();
        ((Action)(() => { _ = range.FirstCol; })).Should().Throw<ObjectDisposedException>();
        ((Action)(() => { _ = range.LastRow; })).Should().Throw<ObjectDisposedException>();
        ((Action)(() => { _ = range.LastCol; })).Should().Throw<ObjectDisposedException>();
    }

    [Fact]
    public void Range_Bounds_Read_Normally_On_A_Live_Workbook()
    {
        using var wb = Workbook.CreateOoxml();
        var range = wb.AddSheet("S").Range("B2:D5");
        range.FirstRow.Should().Be(2);
        range.FirstCol.Should().Be(2);
        range.LastRow.Should().Be(5);
        range.LastCol.Should().Be(4);
    }

    // ---- IRange.Value scalar parity ------------------------------------

    [Fact]
    public void Value_With_Unsupported_Type_Throws_The_Pinned_Scalar_Message()
    {
        using var wb = Workbook.CreateOoxml();
        var range = wb.AddSheet("S").Range("A1:A2");
        ((Action)(() => range.Value(new object())))
            .Should().Throw<ArgumentException>()
            .WithMessage("*is not a supported scalar*");
    }

    [Fact]
    public void Value_Accepts_The_Unsigned_Scalars()
    {
        // ushort/uint/ulong are part of the NPOI engine's dispatch set.
        using var wb = Workbook.CreateOoxml();
        var sheet = wb.AddSheet("S");
        sheet.Range("A1:A1").Value((ushort)7);
        sheet.Range("B1:B1").Value((uint)70_000);
        sheet.Range("C1:C1").Value((ulong)7_000_000_000);

        sheet["A1"].GetNumber().Should().Be(7);
        sheet["B1"].GetNumber().Should().Be(70_000);
        sheet["C1"].GetNumber().Should().Be(7_000_000_000);
    }
}
