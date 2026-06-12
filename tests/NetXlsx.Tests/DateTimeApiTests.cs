// Coverage for the date/time/duration surface (v0.3.x date-time slice).

using System;
using AwesomeAssertions;
using Xunit;

namespace NetXlsx.Tests;

public class DateTimeApiTests
{
    [Fact]
    public void SetDate_DateTime_Roundtrips_Verbatim()
    {
        var dt = new DateTime(2026, 5, 16, 9, 30, 15, DateTimeKind.Unspecified);
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetDate(dt);

        sheet["A1"].Kind.Should().Be(CellKind.Date);
        sheet["A1"].GetDate().Should().Be(dt);
        // Decision I17: result Kind is always Unspecified.
        sheet["A1"].GetDate()!.Value.Kind.Should().Be(DateTimeKind.Unspecified);
    }

    [Fact]
    public void SetDate_Kind_Is_Stored_As_Is_No_Conversion()
    {
        // Decision I17: DateTime.Kind on write is ignored — no timezone
        // conversion. On read, Kind is always Unspecified.
        var local = new DateTime(2026, 5, 16, 9, 30, 0, DateTimeKind.Local);
        var utc = new DateTime(2026, 5, 16, 9, 30, 0, DateTimeKind.Utc);

        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetDate(local);
        sheet["A2"].SetDate(utc);

        // Both round-trip to the same wall-clock instant (no conversion).
        sheet["A1"].GetDate()!.Value.Ticks.Should().Be(local.Ticks);
        sheet["A2"].GetDate()!.Value.Ticks.Should().Be(utc.Ticks);
        sheet["A1"].GetDate()!.Value.Kind.Should().Be(DateTimeKind.Unspecified);
    }

    [Fact]
    public void SetDate_DateOnly_Roundtrips()
    {
        var d = new DateOnly(2026, 5, 16);
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetDate(d);

        sheet["A1"].Kind.Should().Be(CellKind.Date);
        sheet["A1"].GetDateOnly().Should().Be(d);
        sheet["A1"].GetDate().Should().Be(d.ToDateTime(TimeOnly.MinValue));
    }

    [Fact]
    public void SetTime_TimeOnly_Roundtrips()
    {
        var t = new TimeOnly(9, 30, 15);
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetTime(t);

        sheet["A1"].GetTime().Should().Be(t);
        // Stored as fraction-of-day, so GetNumber should give back ~0.396...
        sheet["A1"].GetNumber()!.Value.Should().BeApproximately(t.ToTimeSpan().TotalDays, 1e-9);
    }

    [Fact]
    public void GetTime_Rounds_15_Digit_Serial_To_Nearest_Millisecond()
    {
        // R-6: LibreOffice/Excel author serials at 15 significant digits;
        // 0.396006944444444 is LO's 9:30:15. Truncation read it back as
        // 9:30:14.9999996 — GetTime must round to the millisecond, agreeing
        // with GetDate/FromSerial and the §7.10 display path on the same cell.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetNumber(0.396006944444444);

        sheet["A1"].GetTime().Should().Be(new TimeOnly(9, 30, 15));
    }

    [Fact]
    public void GetDuration_Rounds_15_Digit_Serial_To_Nearest_Millisecond()
    {
        // 27.5 hours + 15 s at LO's 15-digit precision: 1.14600694444444 days
        // is 27:30:15 minus ~0.3 µs — must read back as exactly 27:30:15.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetNumber(1.14600694444444);

        sheet["A1"].GetDuration().Should().Be(new TimeSpan(27, 30, 15));
    }

    [Fact]
    public void GetTime_Returns_Null_When_Serial_Rounds_Up_To_Midnight()
    {
        // R-6 boundary pin: a serial inside [0, 1) that ROUNDS to exactly
        // 24:00:00 is unrepresentable in TimeOnly — the contract's "null when
        // outside [0, 1)" applies to the post-rounding value, not 00:00:00.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetNumber(0.99999999999); // < 1.0, but < half a ms from 24h

        sheet["A1"].GetTime().Should().BeNull();
        // The same serial as a duration is fine — TimeSpan can hold 24h.
        sheet["A1"].GetDuration().Should().Be(TimeSpan.FromDays(1));
    }

    [Fact]
    public void GetTime_Returns_Null_For_Values_Outside_TimeOnly_Range()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetNumber(1.5);   // > 1 day — not a valid time-of-day
        sheet["A2"].SetNumber(-0.25); // negative

        sheet["A1"].GetTime().Should().BeNull();
        sheet["A2"].GetTime().Should().BeNull();
    }

    [Fact]
    public void SetDuration_TimeSpan_Roundtrips()
    {
        var d = TimeSpan.FromHours(25.5);  // exceeds 24h — valid duration
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetDuration(d);

        sheet["A1"].GetDuration()!.Value.Should().BeCloseTo(d, TimeSpan.FromMicroseconds(1));
    }

    [Fact]
    public void SetDuration_Throws_On_Negative_TimeSpan()
    {
        // Decision I15.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        Action act = () => sheet["A1"].SetDuration(TimeSpan.FromHours(-1));
        act.Should().Throw<ArgumentOutOfRangeException>()
            .WithMessage("*Negative TimeSpan*Excel*decision I15*");
    }

    [Fact]
    public void GetDate_Returns_Null_For_Non_Date_Cells()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        sheet["A1"].SetString("hello");
        sheet["A2"].SetNumber(42);       // numeric, not date-formatted
        sheet["A3"].SetBool(true);

        sheet["A1"].GetDate().Should().BeNull();
        sheet["A2"].GetDate().Should().BeNull();
        sheet["A3"].GetDate().Should().BeNull();
    }

    [Fact]
    public void IRow_Set_Has_Fluent_Overloads_For_All_DateTime_Types()
    {
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var dt = new DateTime(2026, 5, 16, 12, 0, 0, DateTimeKind.Unspecified);
        var d = new DateOnly(2026, 5, 16);
        var t = new TimeOnly(12, 30, 0);
        var span = TimeSpan.FromHours(3.5);

        sheet.AppendRow()
            .Set(1, dt)
            .Set(2, d)
            .Set(3, t)
            .Set(4, span);

        sheet["A1"].GetDate().Should().Be(dt);
        sheet["B1"].GetDateOnly().Should().Be(d);
        sheet["C1"].GetTime().Should().Be(t);
        sheet["D1"].GetDuration()!.Value.Should().BeCloseTo(span, TimeSpan.FromMicroseconds(1));
    }

    [Fact]
    public void SetDate_Preserves_Existing_Explicit_Style()
    {
        // Decision I-18: default date format applies only if cell has no
        // prior style — exercised through the public styling API.
        using var wb = Workbook.Create();
        var sheet = wb.AddSheet("S");
        var cell = sheet["A1"];

        // Set a non-default date format before writing the date.
        cell.NumberFormat("dd-mmm-yyyy");

        cell.SetDate(new DateOnly(2026, 5, 16));

        // The custom format should be preserved (not replaced by our default).
        cell.GetStyle().NumberFormat.Should().Be("dd-mmm-yyyy");
    }
}
