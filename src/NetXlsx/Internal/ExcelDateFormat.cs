// I-82 engine swap — Excel date/time number-format renderer (closeout slice).
//
// Renders a date/time number-format code applied to a date cell's serial for
// ICell.GetString on the SDK engine (design §7.10: only date-formatted cells
// are display-formatted; plain numbers stay invariant G17). The NPOI engine
// delegates the same job to NPOI's DataFormatter; this renderer reproduces the
// agreed behavior over the cross-engine matrix and is Excel-correct where NPOI
// is demonstrably not (oracle-dumped 2026-06-03, pinned in
// DateGetStringTests + CrossEngineDifferentialTests):
//   - quoted literals: "yyyy\"y\"" renders 2026y (NPOI re-interprets quoted
//     content and keeps the quote characters — a mangle, not a contract);
//   - lowercase meridiems: am/pm + a/p render am/pm + a/p (NPOI emits "a6/p6"
//     or falls back to the raw serial depending on token position);
//   - a meridiem anywhere in the code switches hours to 12-hour rendering
//     (NPOI only recognizes the exact uppercase "AM/PM" / "A/P" after the
//     hour field).
// Both engines fall back to the raw G17 serial for negative serials (Excel
// renders ###### — there is no date to show; the raw value is the honest
// fallback both engines agree on).
//
// Output is culture-independent: NPOI's DataFormatter emits invariant-English
// month/day names regardless of WorkbookOptions.DisplayCulture (oracle-pinned
// under de-DE), so this renderer uses CultureInfo.InvariantCulture names.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace NetXlsx;

internal static class ExcelDateFormat
{
    /// <summary>
    /// Renders <paramref name="serial"/> through an Excel date/time format
    /// code. <paramref name="dt"/> is the epoch-resolved, millisecond-rounded
    /// DateTime for the serial (the workbook's FromSerial); the raw serial
    /// feeds elapsed-time tokens ([h]/[m]/[s]). Returns null when the code
    /// contains no date/time fields (caller falls back to invariant G17).
    /// </summary>
    internal static string? TryFormat(string code, DateTime dt, double serial)
    {
        var tokens = Tokenize(FirstSection(code));
        DisambiguateMinutes(tokens);
        bool twelveHour = tokens.Exists(t => t.Kind == TokenKind.Meridiem);
        bool anyField = tokens.Exists(t =>
            t.Kind is TokenKind.Field or TokenKind.Elapsed or TokenKind.Meridiem);
        if (!anyField) return null;

        var sb = new StringBuilder();
        foreach (var t in tokens)
            Render(sb, t, dt, serial, twelveHour);
        return sb.ToString();
    }

    // ---- Tokenizer -----------------------------------------------------------

    private enum TokenKind { Literal, Field, Elapsed, Meridiem, Fraction }

    // Field/Elapsed letter is lowercase y/m/d/h/s; minute disambiguation
    // rewrites month 'm' to 'M' (month) vs 'm' (minute) in a second pass.
    private readonly record struct Token(TokenKind Kind, char Letter, int Count, string? Text);

    // Excel format sections are <positive>;<negative>;<zero>;<text> — a date
    // value always renders through the first section (oracle: "m/d/yyyy;@").
    private static string FirstSection(string code)
    {
        bool inQuote = false;
        for (int i = 0; i < code.Length; i++)
        {
            char ch = code[i];
            if (ch == '"') inQuote = !inQuote;
            else if (ch == '\\') i++;
            else if (ch == ';' && !inQuote) return code.Substring(0, i);
        }
        return code;
    }

    private static List<Token> Tokenize(string section)
    {
        var tokens = new List<Token>();
        var literal = new StringBuilder();

        void FlushLiteral()
        {
            if (literal.Length == 0) return;
            tokens.Add(new Token(TokenKind.Literal, '\0', 0, literal.ToString()));
            literal.Clear();
        }

        for (int i = 0; i < section.Length; i++)
        {
            char ch = section[i];
            switch (ch)
            {
                case '"':
                {
                    // Quoted literal — content renders verbatim (Excel-correct;
                    // see the header note on the NPOI divergence).
                    int close = section.IndexOf('"', i + 1);
                    if (close < 0) { literal.Append(section, i + 1, section.Length - i - 1); i = section.Length; }
                    else { literal.Append(section, i + 1, close - i - 1); i = close; }
                    break;
                }
                case '\\':
                    if (i + 1 < section.Length) literal.Append(section[++i]);
                    break;
                case '_':
                    // Skip-width: renders as a space the width of the next char.
                    literal.Append(' ');
                    if (i + 1 < section.Length) i++;
                    break;
                case '*':
                    // Fill-repeat: meaningless outside a real cell width — drop
                    // the marker and its fill char.
                    if (i + 1 < section.Length) i++;
                    break;
                case '[':
                {
                    int close = section.IndexOf(']', i);
                    if (close < 0) { i = section.Length; break; }
                    string inner = section.Substring(i + 1, close - i - 1);
                    char c0 = inner.Length > 0 ? char.ToLowerInvariant(inner[0]) : '\0';
                    bool elapsed = (c0 == 'h' || c0 == 'm' || c0 == 's') && IsRunOf(inner, inner[0]);
                    if (elapsed)
                    {
                        FlushLiteral();
                        tokens.Add(new Token(TokenKind.Elapsed, c0, inner.Length, null));
                    }
                    // else: [Red] / [$-409] / [>=100] — stripped.
                    i = close;
                    break;
                }
                case '.':
                {
                    // A '.' starts a fractional-seconds field only when followed
                    // by '0' AND the last field token is seconds-like; otherwise
                    // it is a literal separator (oracle: "dd.mm.yyyy").
                    int zeros = 0;
                    while (i + 1 + zeros < section.Length && section[i + 1 + zeros] == '0') zeros++;
                    if (zeros > 0 && LastFieldIsSeconds(tokens))
                    {
                        FlushLiteral();
                        tokens.Add(new Token(TokenKind.Fraction, '\0', zeros, null));
                        i += zeros;
                    }
                    else
                    {
                        literal.Append('.');
                    }
                    break;
                }
                default:
                {
                    // Meridiem sequences (case-insensitive): AM/PM, A/P.
                    if (MatchMeridiem(section, i, out int len, out bool longForm, out bool lower))
                    {
                        FlushLiteral();
                        tokens.Add(new Token(TokenKind.Meridiem, longForm ? 'l' : 's', 0, lower ? "lower" : "upper"));
                        i += len - 1;
                        break;
                    }
                    char low = char.ToLowerInvariant(ch);
                    if (low is 'y' or 'm' or 'd' or 'h' or 's')
                    {
                        int count = 1;
                        while (i + count < section.Length && char.ToLowerInvariant(section[i + count]) == low) count++;
                        FlushLiteral();
                        tokens.Add(new Token(TokenKind.Field, low, count, null));
                        i += count - 1;
                    }
                    else
                    {
                        literal.Append(ch);
                    }
                    break;
                }
            }
        }
        FlushLiteral();
        return tokens;
    }

    private static bool IsRunOf(string s, char first)
    {
        foreach (char c in s)
            if (char.ToLowerInvariant(c) != char.ToLowerInvariant(first)) return false;
        return s.Length > 0;
    }

    private static bool LastFieldIsSeconds(List<Token> tokens)
    {
        for (int i = tokens.Count - 1; i >= 0; i--)
        {
            var t = tokens[i];
            if (t.Kind == TokenKind.Field || t.Kind == TokenKind.Elapsed)
                return t.Letter == 's';
        }
        return false;
    }

    private static bool MatchMeridiem(string s, int i, out int len, out bool longForm, out bool lower)
    {
        if (StartsWithIgnoreCase(s, i, "AM/PM"))
        {
            len = 5; longForm = true;
            lower = s[i] == 'a';
            return true;
        }
        if (StartsWithIgnoreCase(s, i, "A/P"))
        {
            len = 3; longForm = false;
            lower = s[i] == 'a';
            return true;
        }
        len = 0; longForm = false; lower = false;
        return false;
    }

    private static bool StartsWithIgnoreCase(string s, int i, string what)
    {
        if (i + what.Length > s.Length) return false;
        for (int k = 0; k < what.Length; k++)
            if (char.ToUpperInvariant(s[i + k]) != what[k]) return false;
        return true;
    }

    // An 'm' run is minutes when the nearest preceding field is hours, or the
    // nearest following field is seconds (Excel's rule; oracle-verified over
    // "m/d/yyyy h:mm", "mm:ss", "yyyy-mm-dd hh:mm:ss"). Runs of 3+ are always
    // months (minute fields cap at two digits). Month keeps letter 'M'.
    private static void DisambiguateMinutes(List<Token> tokens)
    {
        for (int i = 0; i < tokens.Count; i++)
        {
            var t = tokens[i];
            if (t.Kind != TokenKind.Field || t.Letter != 'm') continue;
            if (t.Count >= 3) { tokens[i] = t with { Letter = 'M' }; continue; }

            char prev = NearestField(tokens, i, -1);
            char next = NearestField(tokens, i, +1);
            bool minute = prev == 'h' || next == 's';
            tokens[i] = t with { Letter = minute ? 'm' : 'M' };
        }
    }

    private static char NearestField(List<Token> tokens, int from, int step)
    {
        for (int i = from + step; i >= 0 && i < tokens.Count; i += step)
        {
            var t = tokens[i];
            if (t.Kind == TokenKind.Field || t.Kind == TokenKind.Elapsed)
                return t.Letter;
        }
        return '\0';
    }

    // ---- Rendering -----------------------------------------------------------

    private static void Render(StringBuilder sb, Token t, DateTime dt, double serial, bool twelveHour)
    {
        var inv = CultureInfo.InvariantCulture;
        switch (t.Kind)
        {
            case TokenKind.Literal:
                sb.Append(t.Text);
                break;

            case TokenKind.Field:
                switch (t.Letter)
                {
                    case 'y':
                        sb.Append(t.Count >= 3
                            ? dt.Year.ToString(inv)
                            : (dt.Year % 100).ToString("00", inv));
                        break;
                    case 'M':
                        sb.Append(t.Count switch
                        {
                            1 => dt.Month.ToString(inv),
                            2 => dt.Month.ToString("00", inv),
                            3 => inv.DateTimeFormat.GetAbbreviatedMonthName(dt.Month),
                            4 => inv.DateTimeFormat.GetMonthName(dt.Month),
                            // mmmmm — first letter of the month name.
                            _ => inv.DateTimeFormat.GetMonthName(dt.Month).Substring(0, 1),
                        });
                        break;
                    case 'd':
                        sb.Append(t.Count switch
                        {
                            1 => dt.Day.ToString(inv),
                            2 => dt.Day.ToString("00", inv),
                            3 => inv.DateTimeFormat.GetAbbreviatedDayName(dt.DayOfWeek),
                            _ => inv.DateTimeFormat.GetDayName(dt.DayOfWeek),
                        });
                        break;
                    case 'h':
                    {
                        int hour = twelveHour ? ((dt.Hour + 11) % 12) + 1 : dt.Hour;
                        sb.Append(t.Count >= 2 ? hour.ToString("00", inv) : hour.ToString(inv));
                        break;
                    }
                    case 'm': // minute (post-disambiguation)
                        sb.Append(t.Count >= 2 ? dt.Minute.ToString("00", inv) : dt.Minute.ToString(inv));
                        break;
                    case 's':
                        sb.Append(t.Count >= 2 ? dt.Second.ToString("00", inv) : dt.Second.ToString(inv));
                        break;
                }
                break;

            case TokenKind.Elapsed:
            {
                // Elapsed tokens read the raw stored serial (epoch offsets and
                // the 1900 phantom day included — NPOI parity, oracle-pinned).
                // Hours/minutes floor; a seconds total rounds half-away
                // (oracle: [s] on …930.5 renders …931 while [h]:mm:ss on the
                // same serial keeps :30). The seconds product is denoised at
                // millisecond precision first — a serial carrying a .5 second
                // multiplies out to …930.4999995, and rounding the raw product
                // would land on the wrong side of the half (cross-engine
                // pinned against NPOI).
                long value = t.Letter switch
                {
                    'h' => (long)(serial * 24.0),
                    'm' => (long)(serial * 1440.0),
                    _ => (long)Math.Round(Math.Round(serial * 86400000.0) / 1000.0, MidpointRounding.AwayFromZero),
                };
                sb.Append(value.ToString(inv).PadLeft(t.Count, '0'));
                break;
            }

            case TokenKind.Meridiem:
            {
                bool pm = dt.Hour >= 12;
                bool lower = t.Text == "lower";
                string text = t.Letter == 'l' ? (pm ? "PM" : "AM") : (pm ? "P" : "A");
                sb.Append(lower ? text.ToLowerInvariant() : text);
                break;
            }

            case TokenKind.Fraction:
            {
                // Millisecond fraction, truncated to the digit count
                // (oracle: "mm:ss.0" on .500 renders ".5").
                int scaled = (int)(dt.Millisecond * Math.Pow(10, t.Count) / 1000.0);
                sb.Append('.').Append(scaled.ToString(inv).PadLeft(t.Count, '0'));
                break;
            }
        }
    }
}
