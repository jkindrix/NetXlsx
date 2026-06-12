; Unshipped analyzer release.
; https://github.com/dotnet/roslyn-analyzers/blob/main/src/Microsoft.CodeAnalysis.Analyzers/ReleaseTrackingAnalyzers.Help.md

### New Rules

Rule ID      | Category   | Severity | Notes
-------------|------------|----------|------
NXLS0001 | NetXlsx | Error    | [Worksheet] type has duplicate [Column] Order values
NXLS0002 | NetXlsx | Error    | [Worksheet] type has no designated constructor
NXLS0003 | NetXlsx | Error    | [Column] Format string failed smoke check
NXLS0004 | NetXlsx | Warning  | [Worksheet] property is neither mapped nor ignored
NXLS0005 | NetXlsx | Error    | [Worksheet] type must be partial
NXLS0006 | NetXlsx | Error    | [Worksheet] property type has no built-in converter
NXLS0007 | NetXlsx | Error    | [Worksheet] type cannot be constructed by the generated ReadRows
