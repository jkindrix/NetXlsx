// I-89 — the standard Office theme, embedded lazily on the first theme-color
// write (the EnsureThemePart choke point in OoxmlWorkbook.Theme.cs, or the
// streaming engine's assembly-time check) and surfaced publicly as
// Workbook.DefaultThemeXml.
//
// Provenance: transcribed VERBATIM (decl + CRLF + single-line body, no
// trailing newline — the exact byte shape Excel writes) from the
// xl/theme/theme1.xml of an Excel-authored workbook, cross-verified
// 2026-06-12 against an independent Word-authored copy of the same theme
// generation; the two agree on every color, format and width value (they
// differ only in localized CJK typeface spellings; this copy carries the
// Excel one). The values are public facts present in every default Office
// document: the current "Office Theme" — clrScheme dk2 #44546A / lt2
// #E7E6E6 / accent1-6 #4472C4 #ED7D31 #A5A5A5 #FFC000 #5B9BD5 #70AD47 /
// hlink #0563C1 / folHlink #954F72, fontScheme Calibri Light / Calibri,
// lnStyleLst widths 6350/12700/19050 EMU. DefaultThemeFactsTests pins all
// of these against the embedded bytes.
//
// Two non-ASCII typeface names (Hang U+B9D1 U+C740 U+0020 U+ACE0 U+B515;
// Hant U+65B0 U+7D30 U+660E U+9AD4) are carried as \uXXXX escapes in regular
// string segments so this source file stays pure ASCII — a transcoding accident
// cannot corrupt them silently (the facts test asserts the decoded names).
//
// The minorFont latin face MUST stay "Calibri": the created-workbook
// stylesheet's font 0 carries <scheme val="minor"/> (OoxmlStylePool), so
// consumers resolve the default cell font through this theme once it is
// embedded — any other minor font would visibly change every cell.

using System.Text;

namespace NetXlsx;

internal static class DefaultTheme
{
    /// <summary>A fresh copy of the default Office theme bytes (the public
    /// Workbook.DefaultThemeXml contract — callers may mutate their copy).</summary>
    internal static byte[] CreateBytes() => (byte[])Bytes.Clone();

    /// <summary>The shared, never-mutated instance for engine-internal reads
    /// (EnsureThemePart / streaming assembly feed the part from a read-only
    /// stream over it). Do not hand this array to callers.</summary>
    internal static byte[] Raw => Bytes;

    private static readonly byte[] Bytes = Encoding.UTF8.GetBytes(Xml);

    private const string Xml =
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""" + "\r\n" +
        """<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">""" +
        """<a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/>""" +
        """</a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/>""" +
        """</a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>""" +
        """<a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>""" +
        """<a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>""" +
        """<a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink>""" +
        """<a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office">""" +
        """<a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/>""" +
        """<a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/>""" +
        "<a:font script=\"Hang\" typeface=\"\uB9D1\uC740 \uACE0\uB515\"/><a:font script=\"Hans\" typeface=\"DengXian Light\"/>" +
        "<a:font script=\"Hant\" typeface=\"\u65B0\u7D30\u660E\u9AD4\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/>" +
        """<a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/>""" +
        """<a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/>""" +
        """<a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/>""" +
        """<a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/>""" +
        """<a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/>""" +
        """<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>""" +
        """<a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/>""" +
        """<a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/>""" +
        """<a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/>""" +
        """<a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/>""" +
        """<a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/>""" +
        """<a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/>""" +
        """<a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/>""" +
        """</a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/>""" +
        """<a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/>""" +
        "<a:font script=\"Hang\" typeface=\"\uB9D1\uC740 \uACE0\uB515\"/><a:font script=\"Hans\" typeface=\"DengXian\"/>" +
        "<a:font script=\"Hant\" typeface=\"\u65B0\u7D30\u660E\u9AD4\"/><a:font script=\"Arab\" typeface=\"Arial\"/>" +
        """<a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/>""" +
        """<a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/>""" +
        """<a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/>""" +
        """<a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/>""" +
        """<a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/>""" +
        """<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>""" +
        """<a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/>""" +
        """<a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/>""" +
        """<a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/>""" +
        """<a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/>""" +
        """<a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/>""" +
        """<a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/>""" +
        """<a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/>""" +
        """</a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill>""" +
        """<a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">""" +
        """<a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/>""" +
        """</a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/>""" +
        """<a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000">""" +
        """<a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/>""" +
        """</a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>""" +
        """<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr">""" +
        """<a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs>""" +
        """<a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/>""" +
        """<a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr">""" +
        """<a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst>""" +
        """<a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst>""" +
        """<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/>""" +
        """</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>""" +
        """<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/>""" +
        """</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>""" +
        """<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/>""" +
        """</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst>""" +
        """<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/>""" +
        """</a:effectStyle><a:effectStyle><a:effectLst>""" +
        """<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">""" +
        """<a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst>""" +
        """</a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/>""" +
        """</a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/>""" +
        """</a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">""" +
        """<a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/>""" +
        """<a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr">""" +
        """<a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/>""" +
        """</a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/>""" +
        """<a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/>""" +
        """</a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/>""" +
        """<a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">""" +
        """<thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/>""" +
        """</a:ext></a:extLst></a:theme>""";

    // (Generated transcription ends. Any edit to the literal above must
    // re-verify against the provenance sources and re-run the facts tests.)
}
