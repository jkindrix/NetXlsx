// Spike 4 — AOT / trim posture
//
// Goal: publish a minimal program that exercises NPOI's xlsx write path
// under PublishAot=true + PublishTrimmed=true. Capture:
//   - Number of trim/AOT analyzer warnings (IL2xxx / IL3xxx)
//   - Does the AOT binary actually run? (i.e., does NPOI's reflection-heavy
//     OOXML serialization survive trimming, or does it crash at runtime?)
//
// Resolves: roadmap pre-impl spike "AOT/trim posture", decision I2.
// Outcome feeds back into the roadmap matrix's TBD* rows for
// "Native AOT compatible" and "Trim compatible".

using System;
using System.IO;
using NPOI.XSSF.UserModel;

namespace NetXlsx.AotSpike;

internal static class Program
{
    private static int Main()
    {
        try
        {
            using var wb = new XSSFWorkbook();
            var sheet = wb.CreateSheet("AotSpike");
            var row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("Hello from AOT");
            row.CreateCell(1).SetCellValue(42.0);
            row.CreateCell(2).SetCellValue(DateTime.UtcNow);

            var tempPath = Path.Combine(Path.GetTempPath(), "netxlsx-aot-spike.xlsx");
            using (var fs = File.Create(tempPath))
            {
                wb.Write(fs);
            }

            var info = new FileInfo(tempPath);
            Console.WriteLine($"OK — wrote {info.Length} bytes to {tempPath}");

            // Round-trip read to exercise the other half of NPOI's surface
            using (var fs = File.OpenRead(tempPath))
            {
                using var read = new XSSFWorkbook(fs);
                var readSheet = read.GetSheetAt(0);
                var readRow = readSheet.GetRow(0);
                Console.WriteLine($"OK — read back '{readRow.GetCell(0).StringCellValue}'");
            }

            File.Delete(tempPath);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"FAILED — {ex.GetType().Name}: {ex.Message}");
            Console.Error.WriteLine(ex.StackTrace);
            return 1;
        }
    }
}
