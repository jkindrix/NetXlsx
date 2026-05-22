// Per-column AutoFilter criteria (decision I-66 / v1.2).
// Models Excel's "custom filter" surface — 1-to-2 conditions joined
// by AND or OR. Each condition is an (operator, value) pair; string
// operators (contains / startsWith / endsWith) encode as the equal
// operator with wildcard-bearing values, matching Excel's filter
// language.

using System;
using System.Collections.Generic;
using System.Globalization;

namespace NetXlsx;

/// <summary>
/// A per-column filter applied to an <see cref="ISheet"/>'s AutoFilter
/// via <see cref="ISheet.SetAutoFilterColumn"/> (decision I-66).
/// Construct via the static factories; combine via
/// <see cref="And"/> / <see cref="Or"/>. Excel allows at most two
/// conditions per column — calling <see cref="And"/> or <see cref="Or"/>
/// on a criteria that already has two conditions throws.
/// </summary>
public sealed class FilterCriteria
{
    /// <summary>Internal: the operator used in OOXML <c>CT_CustomFilter.operator</c>.</summary>
    internal enum Op
    {
        Equal,
        NotEqual,
        GreaterThan,
        GreaterThanOrEqual,
        LessThan,
        LessThanOrEqual,
    }

    /// <summary>Internal: how two conditions join.</summary>
    internal enum Combinator
    {
        Single,
        And,
        Or,
    }

    internal readonly Op Op1;
    internal readonly string Val1;
    internal readonly Op? Op2;
    internal readonly string? Val2;
    internal readonly Combinator Combine;

    private FilterCriteria(Op op1, string val1)
    {
        Op1 = op1;
        Val1 = val1;
        Op2 = null;
        Val2 = null;
        Combine = Combinator.Single;
    }

    private FilterCriteria(Op op1, string val1, Op op2, string val2, Combinator combine)
    {
        Op1 = op1;
        Val1 = val1;
        Op2 = op2;
        Val2 = val2;
        Combine = combine;
    }

    // ---- Equality + ordering (numeric or string) ---------------------

    /// <summary>Equal to <paramref name="value"/>.</summary>
    public static FilterCriteria EqualTo(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return new FilterCriteria(Op.Equal, value);
    }

    /// <summary>Not equal to <paramref name="value"/>.</summary>
    public static FilterCriteria NotEqualTo(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return new FilterCriteria(Op.NotEqual, value);
    }

    /// <summary>Strictly greater than <paramref name="value"/>.</summary>
    public static FilterCriteria GreaterThan(double value) =>
        new FilterCriteria(Op.GreaterThan, value.ToString("R", CultureInfo.InvariantCulture));

    /// <summary>Greater than or equal to <paramref name="value"/>.</summary>
    public static FilterCriteria GreaterThanOrEqual(double value) =>
        new FilterCriteria(Op.GreaterThanOrEqual, value.ToString("R", CultureInfo.InvariantCulture));

    /// <summary>Strictly less than <paramref name="value"/>.</summary>
    public static FilterCriteria LessThan(double value) =>
        new FilterCriteria(Op.LessThan, value.ToString("R", CultureInfo.InvariantCulture));

    /// <summary>Less than or equal to <paramref name="value"/>.</summary>
    public static FilterCriteria LessThanOrEqual(double value) =>
        new FilterCriteria(Op.LessThanOrEqual, value.ToString("R", CultureInfo.InvariantCulture));

    /// <summary>
    /// Numeric range <c>[min, max]</c> — equivalent to
    /// <c>GreaterThanOrEqual(min).And(LessThanOrEqual(max))</c>.
    /// </summary>
    public static FilterCriteria Between(double min, double max)
    {
        return GreaterThanOrEqual(min).And(LessThanOrEqual(max));
    }

    // ---- String-pattern (encoded as equal with wildcards) ------------

    /// <summary>
    /// Matches cells whose string value contains <paramref name="substring"/>
    /// (case-insensitive in Excel). Encoded as equal-to <c>*substring*</c>
    /// via Excel's filter wildcard syntax.
    /// </summary>
    public static FilterCriteria Contains(string substring)
    {
        ArgumentNullException.ThrowIfNull(substring);
        return new FilterCriteria(Op.Equal, "*" + EscapeFilterWildcards(substring) + "*");
    }

    /// <summary>Matches cells whose string value does NOT contain <paramref name="substring"/>.</summary>
    public static FilterCriteria DoesNotContain(string substring)
    {
        ArgumentNullException.ThrowIfNull(substring);
        return new FilterCriteria(Op.NotEqual, "*" + EscapeFilterWildcards(substring) + "*");
    }

    /// <summary>Matches cells whose string value starts with <paramref name="prefix"/>.</summary>
    public static FilterCriteria StartsWith(string prefix)
    {
        ArgumentNullException.ThrowIfNull(prefix);
        return new FilterCriteria(Op.Equal, EscapeFilterWildcards(prefix) + "*");
    }

    /// <summary>Matches cells whose string value ends with <paramref name="suffix"/>.</summary>
    public static FilterCriteria EndsWith(string suffix)
    {
        ArgumentNullException.ThrowIfNull(suffix);
        return new FilterCriteria(Op.Equal, "*" + EscapeFilterWildcards(suffix));
    }

    // ---- Combinators -------------------------------------------------

    /// <summary>
    /// Combines this criteria with <paramref name="other"/> via AND.
    /// Both must be single-condition (Excel allows at most two
    /// conditions per column).
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="other"/> is null.</exception>
    /// <exception cref="InvalidOperationException">Either criteria already has two conditions.</exception>
    public FilterCriteria And(FilterCriteria other)
    {
        ArgumentNullException.ThrowIfNull(other);
        EnsureSingle(this, other);
        return new FilterCriteria(Op1, Val1, other.Op1, other.Val1, Combinator.And);
    }

    /// <summary>Combines this criteria with <paramref name="other"/> via OR (otherwise see <see cref="And"/>).</summary>
    public FilterCriteria Or(FilterCriteria other)
    {
        ArgumentNullException.ThrowIfNull(other);
        EnsureSingle(this, other);
        return new FilterCriteria(Op1, Val1, other.Op1, other.Val1, Combinator.Or);
    }

    private static void EnsureSingle(FilterCriteria a, FilterCriteria b)
    {
        if (a.Combine != Combinator.Single || b.Combine != Combinator.Single)
        {
            throw new InvalidOperationException(
                "Excel allows at most two conditions per filter column. " +
                "Combine two single-condition criteria; do not chain further.");
        }
    }

    // ---- Wildcard escaping -------------------------------------------

    /// <summary>
    /// Escapes Excel's filter wildcards (<c>*</c> and <c>?</c>) by
    /// prefixing them with <c>~</c> — Excel's escape character for
    /// literal wildcards inside filter values.
    /// </summary>
    private static string EscapeFilterWildcards(string s)
    {
        if (s.Length == 0) return s;
        // Fast-path: no wildcards present.
        if (s.IndexOfAny(s_wildcards) < 0) return s;
        var sb = new System.Text.StringBuilder(s.Length + 4);
        foreach (var c in s)
        {
            if (c == '*' || c == '?' || c == '~') sb.Append('~');
            sb.Append(c);
        }
        return sb.ToString();
    }

    private static readonly char[] s_wildcards = new[] { '*', '?', '~' };
}
