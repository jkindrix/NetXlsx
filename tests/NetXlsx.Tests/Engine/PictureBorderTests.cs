// Picture borders — decision I-86.
//
// IPicture.Border is the first mutating IPicture member: a solid line border
// written as <a:ln><a:solidFill>…</a:solidFill></a:ln> on xdr:pic/xdr:spPr.
// The ground truth is the ANIMAL PSS fidelity pin (O-2): the 10306 blister-card
// picture carries exactly <a:ln><a:solidFill><a:schemeClr val="tx1"/>
// </a:solidFill></a:ln> — the read-back test injects that markup verbatim and
// expects the I-81 ALIAS slot path (tx1 → dk1 → index 1) to resolve it.
//
// Pinned I-86 decisions exercised here:
//   - Set is a WHOLESALE replacement of <a:ln>: unmodeled line props (dash,
//     gradient, caps) do not survive a read-modify-write; set-null removes ANY
//     <a:ln> including a non-solid one.
//   - Get returns null for non-representable borders (noFill — the borderless
//     <a:ln w="1"><a:noFill/> idiom renders identically to no <a:ln> at all —
//     gradients, non-srgb/scheme color models, unmapped scheme names like
//     phClr, and color-transform children) rather than a silent approximation.
//   - ThemeColor wins over Color when both are set (the I-79 precedence rule);
//     ThemeColor.Tint must be 0 (drawingML lines carry no cell-style tint axis).
//   - A corrupt a:srgbClr/@val fails loud as MalformedFileException (I-83 /
//     quirk #13 — a new opened-file leaf-text parse site).

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using AwesomeAssertions;
using NetXlsx;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace NetXlsx.Tests.Engine;

public class PictureBorderTests
{
    // Known-valid 1×1 transparent PNG (same fixture as PictureTests).
    private static readonly byte[] OnePixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

    private const string DmlNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static string TempXlsxPath()
        => Path.Combine(Path.GetTempPath(), $"netxlsx-pic-border-{Guid.NewGuid():N}.xlsx");

    /// <summary>The persisted a:ln of the (single) picture in the saved workbook.</summary>
    private static XElement? SavedLine(IWorkbook wb)
        => SavedOoxml.DrawingXml(wb).Root!
            .Descendants(SavedOoxml.Xdr + "pic").Single()
            .Element(SavedOoxml.Xdr + "spPr")!
            .Element(SavedOoxml.Dml + "ln");

    /// <summary>Replaces the picture's a:ln with hand-authored markup via the hatch.</summary>
    private static void InjectLine(IPicture pic, string lnInnerXml)
    {
        var spPr = pic.Underlying.ShapeProperties!;
        spPr.RemoveAllChildren<A.Outline>();
        spPr.Append(new A.Outline($"<a:ln xmlns:a=\"{DmlNs}\">{lnInnerXml}</a:ln>"));
    }

    // ---- Defaults ---------------------------------------------------------

    [Fact]
    public void Border_Defaults_To_Null_On_A_Fresh_Picture()
    {
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);

        pic.Border.Should().BeNull("a freshly added picture has no <a:ln> element");
    }

    // ---- Emission: the ANIMAL PSS shape ------------------------------------

    [Fact]
    public void Border_ThemeColor_Emits_The_AnimalPss_Solid_SchemeClr_Shape()
    {
        // The fidelity pin's target markup is
        //   <a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>;
        // tx1 is the dk1 alias (I-81), so index 1 emits the canonical dk1 slot —
        // render-identical in Excel (spreadsheet drawings carry no clrMap).
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);

        pic.Border = new PictureBorder { ThemeColor = new ThemeColor(1) };

        var ln = SavedLine(wb);
        ln.Should().NotBeNull();
        ln!.Attributes().Where(a => !a.IsNamespaceDeclaration)
            .Should().BeEmpty("no width was requested, so @w must be omitted");
        var solidFill = ln.Elements().Should().ContainSingle().Subject;
        solidFill.Name.Should().Be(SavedOoxml.Dml + "solidFill");
        var scheme = solidFill.Elements().Should().ContainSingle().Subject;
        scheme.Name.Should().Be(SavedOoxml.Dml + "schemeClr");
        scheme.Attributes().Where(a => !a.IsNamespaceDeclaration).Should().ContainSingle()
            .Which.Should().Match<XAttribute>(a => a.Name == "val" && a.Value == "dk1");
        scheme.Elements().Should().BeEmpty("a bare scheme color carries no transform children");

        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Border_Reads_The_AnimalPss_Tx1_Alias_Form()
    {
        // The exact source markup from the ANIMAL PSS drawing4.xml — tx1
        // exercises the I-81 alias slot path (tx1 → dk1 → index 1).
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);
        InjectLine(pic, "<a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill>");

        var border = pic.Border;

        border.Should().NotBeNull();
        border!.ThemeColor.Should().Be(new ThemeColor(1));
        border.Color.Should().BeNull();
        border.WidthPoints.Should().BeNull();
    }

    // ---- Emission: explicit color + width ----------------------------------

    [Fact]
    public void Border_Color_Emits_SrgbClr_And_Reads_Back()
    {
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);

        pic.Border = new PictureBorder { Color = Color.Red };

        var ln = SavedLine(wb)!;
        var srgb = ln.Element(SavedOoxml.Dml + "solidFill")!.Element(SavedOoxml.Dml + "srgbClr");
        srgb.Should().NotBeNull();
        srgb!.Attribute("val")!.Value.Should().Be("FF0000");

        pic.Border.Should().Be(new PictureBorder { Color = Color.Red });
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Border_WidthPoints_Emits_Emu_And_Reads_Back()
    {
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);

        pic.Border = new PictureBorder { Color = Color.Black, WidthPoints = 2.25 };

        var ln = SavedLine(wb)!;
        ln.Attribute("w")!.Value.Should().Be("28575", "2.25 pt × 12700 EMU/pt");

        pic.Border.Should().Be(new PictureBorder { Color = Color.Black, WidthPoints = 2.25 });
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Border_ThemeColor_Wins_Over_Color_When_Both_Set()
    {
        // The I-79 precedence rule, verbatim from CellStyle.BackgroundTheme.
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);

        pic.Border = new PictureBorder { Color = Color.Red, ThemeColor = new ThemeColor(4) };

        var solidFill = SavedLine(wb)!.Element(SavedOoxml.Dml + "solidFill")!;
        solidFill.Element(SavedOoxml.Dml + "schemeClr")!.Attribute("val")!.Value.Should().Be("accent1");
        solidFill.Element(SavedOoxml.Dml + "srgbClr").Should().BeNull();
    }

    [Theory]
    [InlineData(0, "lt1")]
    [InlineData(3, "dk2")]
    [InlineData(9, "accent6")]
    [InlineData(10, "hlink")]
    [InlineData(11, "folHlink")]
    public void Border_ThemeColor_Maps_Every_I81_Slot(int index, string expected)
    {
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);

        pic.Border = new PictureBorder { ThemeColor = new ThemeColor(index) };

        SavedLine(wb)!.Element(SavedOoxml.Dml + "solidFill")!
            .Element(SavedOoxml.Dml + "schemeClr")!.Attribute("val")!.Value.Should().Be(expected);
        pic.Border.Should().Be(new PictureBorder { ThemeColor = new ThemeColor(index) },
            "the slot must round-trip through the read-back");
    }

    // ---- Removal + wholesale replacement -----------------------------------

    [Fact]
    public void Border_Set_Null_Removes_The_Line()
    {
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);
        pic.Border = new PictureBorder { Color = Color.Black };

        pic.Border = null;

        pic.Border.Should().BeNull();
        SavedLine(wb).Should().BeNull("set-null must REMOVE <a:ln>, not leave an empty element");
        OpenXmlValidationGate.AssertValid(wb);
    }

    [Fact]
    public void Border_Set_Null_Removes_A_NonSolid_Line_Too()
    {
        // Pinned I-86 decision: set-null is explicit removal intent and takes
        // ANY <a:ln> with it, including the noFill idiom.
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);
        InjectLine(pic, "<a:noFill/>");

        pic.Border = null;

        SavedLine(wb).Should().BeNull();
    }

    [Fact]
    public void Border_Set_Is_A_Wholesale_Replacement()
    {
        // Unmodeled line props (here @cap and a dash style) must NOT survive —
        // the documented contract is replace, not merge.
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);
        InjectLine(pic,
            "<a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill><a:prstDash val=\"dash\"/>");

        pic.Border = new PictureBorder { Color = Color.Black, WidthPoints = 1 };

        var ln = SavedLine(wb)!;
        ln.Element(SavedOoxml.Dml + "prstDash").Should().BeNull("Set replaces <a:ln> wholesale");
        ln.Element(SavedOoxml.Dml + "solidFill")!
            .Element(SavedOoxml.Dml + "srgbClr")!.Attribute("val")!.Value.Should().Be("000000");
        OpenXmlValidationGate.AssertValid(wb);
    }

    // ---- Non-representable borders read as null -----------------------------

    [Theory]
    [InlineData("<a:noFill/>", "the borderless a:ln + a:noFill idiom")]
    [InlineData("<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/></a:gs>" +
                "<a:gs pos=\"100000\"><a:srgbClr val=\"0000FF\"/></a:gs></a:gsLst>" +
                "<a:lin ang=\"0\" scaled=\"1\"/></a:gradFill>", "a gradient line fill")]
    [InlineData("<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>",
        "a scheme name outside the I-81 slot map")]
    [InlineData("<a:solidFill><a:schemeClr val=\"dk1\"><a:lumMod val=\"75000\"/></a:schemeClr></a:solidFill>",
        "a color-transform child on the scheme color")]
    [InlineData("<a:solidFill><a:srgbClr val=\"FF0000\"><a:alpha val=\"50000\"/></a:srgbClr></a:solidFill>",
        "a color-transform child on the explicit color")]
    [InlineData("<a:solidFill><a:sysClr val=\"windowText\"/></a:solidFill>",
        "a color model the record does not represent")]
    public void Border_Reads_Null_For_NonRepresentable_Lines(string lnInnerXml, string why)
    {
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);
        InjectLine(pic, lnInnerXml);

        pic.Border.Should().BeNull($"reading {why} as a PictureBorder would misreport the rendering");
    }

    [Fact]
    public void Border_Corrupt_SrgbClr_Fails_Loud()
    {
        // A new opened-file leaf-text parse site defaults to fail-loud
        // (I-83 / SDK-quirk #13) — never a silently substituted color.
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);
        InjectLine(pic, "<a:solidFill><a:srgbClr val=\"XYZ\"/></a:solidFill>");

        Action act = () => { var _ = pic.Border; };
        act.Should().Throw<MalformedFileException>().Which.Message.Should().Contain("XYZ");
    }

    // ---- Validation ---------------------------------------------------------

    [Fact]
    public void Border_Rejects_Neither_Color_Axis_Set()
    {
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);

        Action act = () => pic.Border = new PictureBorder { WidthPoints = 1 };
        act.Should().Throw<ArgumentException>().Which.Message.Should().Contain("Color or ThemeColor");
    }

    [Fact]
    public void Border_Rejects_NonZero_Tint()
    {
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);

        Action act = () => pic.Border = new PictureBorder { ThemeColor = new ThemeColor(1, Tint: 0.4) };
        act.Should().Throw<ArgumentException>().Which.Message.Should().Contain("Tint");
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(12)]
    public void Border_Rejects_Theme_Index_Outside_The_Slot_Map(int index)
    {
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);

        Action act = () => pic.Border = new PictureBorder { ThemeColor = new ThemeColor(index) };
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Theory]
    [InlineData(0.0)]
    [InlineData(-1.0)]
    [InlineData(1584.01)]
    [InlineData(double.NaN)]
    public void Border_Rejects_Invalid_WidthPoints(double width)
    {
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);

        Action act = () => pic.Border = new PictureBorder { Color = Color.Black, WidthPoints = width };
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Border_Accepts_The_ST_LineWidth_Maximum()
    {
        using var wb = Workbook.Create();
        var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);

        pic.Border = new PictureBorder { Color = Color.Black, WidthPoints = 1584 };

        SavedLine(wb)!.Attribute("w")!.Value.Should().Be("20116800", "the ST_LineWidth maximum");
        OpenXmlValidationGate.AssertValid(wb);
    }

    // ---- Round-trip + open-mutate -------------------------------------------

    [Fact]
    public void Border_RoundTrips_Through_Save_And_Reopen()
    {
        string path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                var sheet = wb.AddSheet("S");
                sheet.AddPicture("B2", OnePixelPng, ImageFormat.Png).Border =
                    new PictureBorder { ThemeColor = new ThemeColor(1), WidthPoints = 1.5 };
                sheet.AddPicture("D2", OnePixelPng, ImageFormat.Png).Border =
                    new PictureBorder { Color = Color.Blue };
                wb.Save(path);
            }

            using (var wb = Workbook.Open(path))
            {
                var pics = wb["S"].Pictures;
                pics.Should().HaveCount(2);
                pics[0].Border.Should().Be(
                    new PictureBorder { ThemeColor = new ThemeColor(1), WidthPoints = 1.5 });
                pics[1].Border.Should().Be(new PictureBorder { Color = Color.Blue });
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Border_On_An_Opened_Files_Picture_Slots_Before_The_Effect_Tail()
    {
        // Open-mutate fixture (the SDK-quirk #8 habit, applied to
        // CT_ShapeProperties): on an opened file whose spPr already carries a
        // later sibling (<a:effectLst>), the new <a:ln> must be inserted BEFORE
        // it — a bare Append would emit out-of-schema-order XML.
        string path = TempXlsxPath();
        try
        {
            using (var wb = Workbook.Create())
            {
                var pic = wb.AddSheet("S").AddPicture("B2", OnePixelPng, ImageFormat.Png);
                pic.Underlying.ShapeProperties!.Append(new A.EffectList());
                wb.Save(path);
            }

            using (var wb = Workbook.Open(path))
            {
                var pic = wb["S"].Pictures.Should().ContainSingle().Subject;
                pic.Border = new PictureBorder { ThemeColor = new ThemeColor(1) };

                var spPr = SavedOoxml.DrawingXml(wb).Root!
                    .Descendants(SavedOoxml.Xdr + "pic").Single()
                    .Element(SavedOoxml.Xdr + "spPr")!;
                var names = spPr.Elements().Select(e => e.Name.LocalName).ToList();
                names.IndexOf("ln").Should().BeLessThan(names.IndexOf("effectLst"),
                    "<a:ln> precedes the effect tail in CT_ShapeProperties");
                OpenXmlValidationGate.AssertValid(wb);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }
}
