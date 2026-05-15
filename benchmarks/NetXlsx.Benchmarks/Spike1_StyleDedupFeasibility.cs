// Spike 1 — Style-dedup feasibility
//
// Question: what is the cost of *not* deduplicating cell styles versus
// deduplicating via a per-workbook style pool? Specifically: file size,
// write time, and the size of the in-memory style table for a workbook
// with N cells styled from a logical palette of M distinct styles.
//
// Resolves: confirms or refutes design decision #4's <10% / <30% target
// for the dedup vs no-dedup overhead.
//
// Run: dotnet run --project benchmarks/NetXlsx.Benchmarks -c Release -- spike-1

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NetXlsx.Benchmarks;

internal static class Spike1_StyleDedupFeasibility
{
    public static int Run()
    {
        Console.WriteLine("Spike 1 — style-dedup feasibility (NPOI 2.7.3, .NET 9)");
        Console.WriteLine();

        // Parameter grid: rows × cols giving total-cell count, with M
        // distinct logical styles cycled through them.
        var scenarios = new (int rows, int cols, int distinctStyles)[]
        {
            (1_000,  10, 10),     // small, typical dedup case
            (10_000, 10, 10),     // medium, typical
            (50_000, 10, 10),     // large, typical
            (10_000, 10, 100),    // medium, more distinct styles
            (10_000, 10, 1_000),  // pathological — many distinct styles
        };

        Console.WriteLine($"{"Mode",-12} {"Cells",10} {"Distinct",10} {"Time (s)",12} {"File (MB)",12} {"Pool size",12}");
        Console.WriteLine(new string('-', 76));

        foreach (var s in scenarios)
        {
            // Variant A — no dedup: a fresh ICellStyle per (logical style, cell) call
            MeasureNoDedup(s.rows, s.cols, s.distinctStyles);
            // Variant B — dedup: one ICellStyle per logical style, reused
            MeasureDedup(s.rows, s.cols, s.distinctStyles);
        }

        return 0;
    }

    private static void MeasureNoDedup(int rows, int cols, int distinctStyles)
    {
        var path = Path.Combine(Path.GetTempPath(), $"spike1-nodedup-{rows}x{cols}-{distinctStyles}.xlsx");
        var sw = Stopwatch.StartNew();
        try
        {
            using var wb = new XSSFWorkbook();
            var sheet = wb.CreateSheet("Data");

            for (int r = 0; r < rows; r++)
            {
                var row = sheet.CreateRow(r);
                for (int c = 0; c < cols; c++)
                {
                    var cell = row.CreateCell(c);
                    cell.SetCellValue(r * cols + c);
                    int logical = (r * cols + c) % distinctStyles;
                    // Create a *fresh* style for each cell, but populated from
                    // the logical palette. NPOI does not dedupe.
                    var style = wb.CreateCellStyle();
                    style.DataFormat = wb.CreateDataFormat().GetFormat(LogicalNumberFormat(logical));
                    cell.CellStyle = style;
                    if (wb.NumCellStyles > 60_000)
                    {
                        // NPOI / OOXML hard cap is ~64k. Bail before it
                        // throws so we can record the partial result.
                        sw.Stop();
                        Console.WriteLine($"{"NoDedup",-12} {rows * cols,10:N0} {distinctStyles,10:N0} {sw.Elapsed.TotalSeconds,12:F2} {"--",12} {wb.NumCellStyles,12:N0}  (cap hit)");
                        return;
                    }
                }
            }
            using var fs = File.Create(path);
            wb.Write(fs);
            sw.Stop();
            var sizeMb = new FileInfo(path).Length / 1024.0 / 1024.0;
            Console.WriteLine($"{"NoDedup",-12} {rows * cols,10:N0} {distinctStyles,10:N0} {sw.Elapsed.TotalSeconds,12:F2} {sizeMb,12:F2} {wb.NumCellStyles,12:N0}");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
            GC.Collect(); GC.WaitForPendingFinalizers();
        }
    }

    private static void MeasureDedup(int rows, int cols, int distinctStyles)
    {
        var path = Path.Combine(Path.GetTempPath(), $"spike1-dedup-{rows}x{cols}-{distinctStyles}.xlsx");
        var sw = Stopwatch.StartNew();
        try
        {
            using var wb = new XSSFWorkbook();
            var sheet = wb.CreateSheet("Data");

            // Pre-create exactly `distinctStyles` styles and reuse them.
            var pool = new ICellStyle[distinctStyles];
            for (int i = 0; i < distinctStyles; i++)
            {
                pool[i] = wb.CreateCellStyle();
                pool[i].DataFormat = wb.CreateDataFormat().GetFormat(LogicalNumberFormat(i));
            }

            for (int r = 0; r < rows; r++)
            {
                var row = sheet.CreateRow(r);
                for (int c = 0; c < cols; c++)
                {
                    var cell = row.CreateCell(c);
                    cell.SetCellValue(r * cols + c);
                    int logical = (r * cols + c) % distinctStyles;
                    cell.CellStyle = pool[logical];
                }
            }
            using var fs = File.Create(path);
            wb.Write(fs);
            sw.Stop();
            var sizeMb = new FileInfo(path).Length / 1024.0 / 1024.0;
            Console.WriteLine($"{"Dedup",-12} {rows * cols,10:N0} {distinctStyles,10:N0} {sw.Elapsed.TotalSeconds,12:F2} {sizeMb,12:F2} {wb.NumCellStyles,12:N0}");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
            GC.Collect(); GC.WaitForPendingFinalizers();
        }
    }

    // 1000 cyclical number formats — enough variety that two different
    // logical indices produce two different formats.
    private static string LogicalNumberFormat(int i) => (i % 4) switch
    {
        0 => "0",
        1 => "0.00",
        2 => "#,##0",
        _ => "$#,##0.00",
    };
}
