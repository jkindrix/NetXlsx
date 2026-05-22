// XssfWorkbook — named-style registry per decision I-57.
// Core class structure is in XssfWorkbook.cs.

using System;
using System.Collections.Generic;

namespace NetXlsx;

internal sealed partial class XssfWorkbook
{
    private Dictionary<string, CellStyle>? _namedStyles;

    private Dictionary<string, CellStyle> NamedStyles =>
        _namedStyles ??= new Dictionary<string, CellStyle>(StringComparer.OrdinalIgnoreCase);

    public void RegisterStyle(string name, CellStyle style)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(style);
        if (name.Length == 0)
            throw new ArgumentException("Style name cannot be empty.", nameof(name));
        NamedStyles[name] = style;
    }

    public CellStyle? GetRegisteredStyle(string name)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        return _namedStyles is not null && _namedStyles.TryGetValue(name, out var s) ? s : null;
    }

    public IReadOnlyCollection<string> RegisteredStyleNames
    {
        get
        {
            ThrowIfDisposed();
            return _namedStyles is null ? Array.Empty<string>() : _namedStyles.Keys;
        }
    }

    /// <summary>
    /// Internal lookup throwing the canonical "no such name" error message.
    /// Shared by ICell.ApplyNamedStyle and IRange.ApplyNamedStyle.
    /// </summary>
    internal CellStyle ResolveNamedStyleOrThrow(string name)
    {
        var style = GetRegisteredStyle(name);
        if (style is null)
        {
            throw new ArgumentException(
                $"No style is registered under '{name}'. " +
                "Use IWorkbook.RegisterStyle before referencing the name.",
                nameof(name));
        }
        return style;
    }
}
