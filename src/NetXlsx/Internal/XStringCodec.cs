// ST_Xstring escape codec (ECMA-376 §22.9.2.19) — decision I-88 / ledger R-3.
//
// XML-1.0-invalid characters in user CONTENT (cell strings, rich-text runs,
// comments) used to pass the setter silently and explode at Save as a raw
// SDK ArgumentException with no cell context — and CR (0x0D) was silently
// destroyed by XmlWriter newline normalization ("C\rD" read back "C<LF>D").
// Excel's own answer is the _xHHHH_ escape convention: encode at the setter,
// decode on read, and the package bytes are exactly what Excel would write.
//
// Encode set: 0x00–0x1F except tab (0x09) and LF (0x0A) — CR (0x0D) IS
// escaped (MS-OI29500 "shall be escaped"; it is XML-valid but normalized
// away by writers and readers alike); U+FFFE/U+FFFF; lone surrogate halves
// (escaped as their own code unit — spec-legal per MS-OI29500, though
// Excel's own emission for that case is unverified). A literal substring
// that the decoder would mistake for an escape gets its leading underscore
// protected as _x005F_ per the convention.
//
// Hex case: emit UPPERCASE (matches Excel); decode case-insensitively
// (matches POI/ClosedXML/openpyxl-style decoders); the leading 'x' must be
// lowercase for a sequence to count as an escape (ClosedXML pins this).

using System;
using System.Globalization;
using System.Text;

namespace NetXlsx;

internal static class XStringCodec
{
    /// <summary>Encodes user text to the ST_Xstring convention. Returns the
    /// same instance when nothing needs escaping (the overwhelmingly common
    /// case — zero allocation).</summary>
    public static string Encode(string value)
    {
        int first = FirstEncodePoint(value);
        if (first < 0) return value;

        var sb = new StringBuilder(value.Length + 8);
        sb.Append(value, 0, first);
        for (int i = first; i < value.Length; i++)
        {
            char ch = value[i];
            if (NeedsEscape(value, i))
            {
                sb.Append("_x").Append(((int)ch).ToString("X4", CultureInfo.InvariantCulture)).Append('_');
            }
            else if (ch == '_' && LooksLikeEscape(value, i))
            {
                // Literal "_xHHHH_" in user text — protect the underscore so
                // the decoder reproduces the literal, not the escaped char.
                sb.Append("_x005F_");
            }
            else
            {
                sb.Append(ch);
            }
        }
        return sb.ToString();
    }

    /// <summary>Decodes well-formed <c>_xHHHH_</c> sequences back to their
    /// characters. Returns the same instance when no escape is present.</summary>
    public static string Decode(string value)
    {
        int idx = value.IndexOf("_x", StringComparison.Ordinal);
        if (idx < 0) return value;

        var sb = new StringBuilder(value.Length);
        int pos = 0;
        while (idx >= 0)
        {
            sb.Append(value, pos, idx - pos);
            if (TryParseEscape(value, idx, out char decoded))
            {
                sb.Append(decoded);
                pos = idx + 7;
            }
            else
            {
                // Not a well-formed escape — the "_x" is literal text.
                sb.Append("_x");
                pos = idx + 2;
            }
            idx = pos < value.Length ? value.IndexOf("_x", pos, StringComparison.Ordinal) : -1;
        }
        sb.Append(value, pos, value.Length - pos);
        return sb.ToString();
    }

    /// <summary>
    /// First index of a character XML 1.0 cannot represent at all, or -1.
    /// The fail-fast surfaces (sheet names, defined names, formulas) use
    /// this instead of escaping — control characters are never meaningful
    /// there and escaping would change reference semantics (I-88 half (a)).
    /// CR is XML-legal, so unlike the Encode set it is NOT included here.
    /// </summary>
    public static int IndexOfXmlInvalid(string value)
    {
        for (int i = 0; i < value.Length; i++)
        {
            char ch = value[i];
            if (ch < 0x20)
            {
                if (ch is not ('\t' or '\n' or '\r')) return i;
                continue;
            }
            if (ch is '\uFFFE' or '\uFFFF') return i;
            if (char.IsHighSurrogate(ch) && (i + 1 >= value.Length || !char.IsLowSurrogate(value[i + 1]))) return i;
            if (char.IsLowSurrogate(ch) && (i == 0 || !char.IsHighSurrogate(value[i - 1]))) return i;
        }
        return -1;
    }

    // The first index that forces an allocation, or -1 when the string is
    // clean (no escape-set char, no literal escape pattern).
    private static int FirstEncodePoint(string value)
    {
        for (int i = 0; i < value.Length; i++)
        {
            if (NeedsEscape(value, i)) return i;
            if (value[i] == '_' && LooksLikeEscape(value, i)) return i;
        }
        return -1;
    }

    private static bool NeedsEscape(string value, int index)
    {
        char ch = value[index];
        if (ch < 0x20) return ch != '\t' && ch != '\n';   // CR included by design
        if (ch is '\uFFFE' or '\uFFFF') return true;
        if (char.IsHighSurrogate(ch))
            return index + 1 >= value.Length || !char.IsLowSurrogate(value[index + 1]);
        if (char.IsLowSurrogate(ch))
            return index == 0 || !char.IsHighSurrogate(value[index - 1]);
        return false;
    }

    // True when value[index..] reads "_xHHHH_" with a lowercase 'x' and four
    // hex digits — i.e. the decoder would treat it as an escape.
    private static bool LooksLikeEscape(string value, int index)
        => TryParseEscape(value, index, out _);

    private static bool TryParseEscape(string value, int index, out char decoded)
    {
        decoded = '\0';
        if (index + 7 > value.Length) return false;
        if (value[index] != '_' || value[index + 1] != 'x' || value[index + 6] != '_') return false;
        int code = 0;
        for (int i = index + 2; i < index + 6; i++)
        {
            int d = HexValue(value[i]);
            if (d < 0) return false;
            code = (code << 4) | d;
        }
        decoded = (char)code;
        return true;
    }

    private static int HexValue(char c) => c switch
    {
        >= '0' and <= '9' => c - '0',
        >= 'a' and <= 'f' => c - 'a' + 10,
        >= 'A' and <= 'F' => c - 'A' + 10,
        _ => -1,
    };
}
