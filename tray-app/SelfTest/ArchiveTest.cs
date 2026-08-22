using System.IO.Compression;
using MediaPipelineTray.Services;

namespace MediaPipelineTray.SelfTest;

/// <summary>
/// Checks collecting a finished job's output into a zip staged for upload.
///
/// The interesting cases are a grouped preset, whose output sits in set folders that must
/// survive into the archive, and output that has since been archived or deleted, which should
/// produce a partial zip with a report rather than an error.
/// </summary>
internal static class ArchiveTest
{
    private static int _failures;

    public static int Run()
    {
        Console.WriteLine("== Output archiving");

        var root = Path.Combine(Path.GetTempPath(), "mp-zip-test-" + Guid.NewGuid().ToString("n")[..8]);

        try
        {
            var paths = new PipelinePaths(root, root);
            var service = new ArchiveService(paths);

            var outputDir = paths.OutputDirectory("sets", "LC");
            Directory.CreateDirectory(Path.Combine(outputDir, "sunny upload"));
            Directory.CreateDirectory(Path.Combine(outputDir, "quiet orchard"));

            // A grouped preset writes into set folders, so the relative paths carry them.
            var relatives = new[]
            {
                @"sunny upload\IMG_0001.JPG",
                @"sunny upload\IMG_0002.JPG",
                @"quiet orchard\IMG_0003.JPG",
            };

            foreach (var relative in relatives)
            {
                File.WriteAllBytes(Path.Combine(outputDir, relative), new byte[2048]);
            }

            var result = service.Create("sets", "LC", relatives, nameHint: "batch.mp4");

            Check("the zip lands in the workspace sync folder",
                result.Path.StartsWith(paths.SyncDirectory("LC"), StringComparison.OrdinalIgnoreCase),
                result.Path);

            Check("it is named after the preset", Path.GetFileName(result.Path).StartsWith("sets-"),
                Path.GetFileName(result.Path));

            Check("every file was collected", result.FileCount == 3, $"got {result.FileCount}");
            Check("nothing was reported missing", result.Missing.Count == 0,
                string.Join(", ", result.Missing));
            Check("the file exists on disk", File.Exists(result.Path), "missing");

            using (var zip = ZipFile.OpenRead(result.Path))
            {
                var names = zip.Entries.Select(e => e.FullName).OrderBy(n => n).ToArray();

                Check("it holds three entries", names.Length == 3, string.Join(", ", names));

                // Set folders must survive, or unpacking would collapse the sets together.
                Check("set folders are preserved",
                    names.Contains("sunny upload/IMG_0001.JPG") && names.Contains("quiet orchard/IMG_0003.JPG"),
                    string.Join(", ", names));

                Check("entries use forward slashes",
                    names.All(n => !n.Contains('\\')), string.Join(", ", names));
            }

            // Output that has since been archived away.
            File.Delete(Path.Combine(outputDir, relatives[1]));
            var partial = service.Create("sets", "LC", relatives);

            Check("a missing file does not fail the archive", partial.FileCount == 2,
                $"got {partial.FileCount}");
            Check("what was skipped is reported", partial.Missing.Count == 1,
                $"got {partial.Missing.Count}");

            // Nothing left at all is a real error, not a zip with no entries.
            foreach (var relative in relatives)
            {
                var path = Path.Combine(outputDir, relative);
                if (File.Exists(path)) { File.Delete(path); }
            }

            var threw = false;
            try
            {
                service.Create("sets", "LC", relatives);
            }
            catch (InvalidOperationException)
            {
                threw = true;
            }

            Check("an empty result is an error, not an empty zip", threw, "no exception");

            var leftovers = Directory.GetFiles(paths.SyncDirectory("LC"), "*.building");
            Check("no half-built archives are left behind", leftovers.Length == 0,
                string.Join(", ", leftovers));
        }
        finally
        {
            try { Directory.Delete(root, recursive: true); } catch (IOException) { }
        }

        Console.WriteLine();
        Console.WriteLine(_failures == 0 ? "ARCHIVING OK" : $"ARCHIVING FAILED: {_failures}");
        return _failures == 0 ? 0 : 1;
    }

    private static void Check(string label, bool passed, string detail)
    {
        if (passed)
        {
            Console.WriteLine($"   PASS  {label}");
            return;
        }

        _failures++;
        Console.WriteLine($"   FAIL  {label}");

        if (detail.Trim().Length > 0)
        {
            Console.WriteLine($"         {detail.Trim()}");
        }
    }
}
