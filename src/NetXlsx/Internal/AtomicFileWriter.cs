// Atomic Save(path) for both engines (remediation R-1/R-2): a failed save must
// never truncate or destroy a pre-existing destination file. The bytes are
// written to a sibling temp file and promoted with a same-volume rename; any
// failure deletes the temp and leaves the destination exactly as it was.
//
// The temp name EMBEDS the destination filename, which is load-bearing for the
// streaming engine: its single-shot finalize must not be burned by a bad path
// (OoxmlStreamingWorkbook.Save), so directory-level errors (missing dir, no
// permission) AND filename-level errors (invalid characters, component too
// long) all surface at temp creation — before the writer callback runs.
// Accepted residual (ledger R-2): destination-is-a-directory and
// destination-locked failures surface at the final File.Move, after the write.

using System;
using System.IO;

namespace NetXlsx;

internal static class AtomicFileWriter
{
    internal static void Write(string path, Action<FileStream> write)
    {
        string fullPath = Path.GetFullPath(path);
        string? dir = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrEmpty(dir)) dir = Path.GetPathRoot(fullPath);
        string tmp = Path.Combine(dir!, $"{Path.GetFileName(fullPath)}.{Guid.NewGuid():N}.netxlsx.tmp");
        try
        {
            using (var fs = new FileStream(tmp, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                write(fs);
            File.Move(tmp, fullPath, overwrite: true);
        }
        catch
        {
            // Best-effort cleanup; the original exception is the one that matters.
            try { File.Delete(tmp); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            throw;
        }
    }
}
