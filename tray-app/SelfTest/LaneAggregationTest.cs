using System.Text;
using System.Text.Json;
using MediaPipelineTray.Models;
using MediaPipelineTray.Services;

namespace MediaPipelineTray.SelfTest;

/// <summary>
/// Checks that concurrent work in one preset is reported as one unit.
///
/// The watcher processes several files at once within a lane, so three images in image-bulk
/// are three jobs in the event stream. This drives the monitor with a synthetic stream rather
/// than racing a real one, so the result does not depend on catching a fast job mid-flight.
/// </summary>
internal static class LaneAggregationTest
{
    private static int _failures;

    public static async Task<int> RunAsync()
    {
        Console.WriteLine("== Lane aggregation");

        var root = Path.Combine(Path.GetTempPath(), "mp-lane-test-" + Guid.NewGuid().ToString("n")[..8]);
        var logs = Path.Combine(root, "logs");
        Directory.CreateDirectory(logs);

        var paths = new PipelinePaths(root, root);

        // The monitor discards in-flight work when the watcher is not running, so the named
        // lock it probes for has to exist. Created unowned deliberately: only its existence
        // matters, and an owned mutex cannot be released across an await.
        using var mutex = new Mutex(false, paths.MutexName, out var createdNew);

        if (!createdNew)
        {
            Console.WriteLine("   SKIP  another instance holds the lock for this root");
            return 0;
        }

        try
        {
            var eventFile = Path.Combine(logs, $"events-{DateTimeOffset.Now:yyyyMMdd}.jsonl");
            var lines = new List<string>();

            // Three files in one preset, started together, plus one elsewhere.
            foreach (var (jobId, file) in new[] { ("aaa", "one.jpg"), ("bbb", "two.jpg"), ("ccc", "three.jpg") })
            {
                lines.Add(Event("job.start", jobId, "image-bulk", "LC", files: [file]));
            }

            lines.Add(Event("job.start", "ddd", "video-clean", "PL", files: ["clip.mp4"]));

            // Differing progress, so the combined figure cannot come from just one job.
            lines.Add(Variant("aaa", "image-bulk", "LC", "one.jpg", 10, 50, "a1.jpg"));
            lines.Add(Variant("bbb", "image-bulk", "LC", "two.jpg", 20, 50, "b1.jpg"));
            lines.Add(Variant("ccc", "image-bulk", "LC", "three.jpg", 30, 50, "c1.jpg"));
            lines.Add(Variant("ddd", "video-clean", "PL", "clip.mp4", 1, 4, "d1.mp4"));

            await File.WriteAllLinesAsync(eventFile, lines, new UTF8Encoding(false));

            var watcher = new WatcherService(paths);
            var monitor = new PipelineMonitor(watcher, new EventStreamReader(paths));

            monitor.PrimeFromToday();
            var snapshot = monitor.Tick();

            Check("two lanes are running, not four jobs", snapshot.Running.Count == 2,
                $"got {snapshot.Running.Count}");

            var bulk = snapshot.Running.FirstOrDefault(lane => lane.Preset == "image-bulk");

            if (bulk is null)
            {
                Check("the image-bulk lane is present", false, "missing");
                return Report();
            }

            Check("it covers all three files", bulk.FileCount == 3, $"got {bulk.FileCount}");
            Check("progress is summed", bulk.VariantsDone == 60, $"got {bulk.VariantsDone}");
            Check("the total is summed", bulk.VariantsTotal == 150, $"got {bulk.VariantsTotal}");
            Check("the fraction is combined", Math.Abs(bulk.Fraction - 0.4) < 0.001,
                bulk.Fraction.ToString("0.000"));
            Check("the subject reads as a count", bulk.Subject == "3 files", bulk.Subject);

            var single = snapshot.Running.FirstOrDefault(lane => lane.Preset == "video-clean");
            Check("a lane with one file names it", single?.Subject == "clip.mp4",
                single?.Subject ?? "missing");

            // Finishing one file must not collapse the whole lane.
            await File.AppendAllLinesAsync(eventFile, [Done("aaa", "image-bulk", "LC", 50)]);

            var after = monitor.Tick();
            var bulkAfter = after.Running.FirstOrDefault(lane => lane.Preset == "image-bulk");

            Check("the lane stays while its other files run", bulkAfter is not null, "lane vanished");
            Check("it now covers two files", bulkAfter?.FileCount == 2, $"got {bulkAfter?.FileCount}");
            Check("the finished file is recorded", after.Finished.Count == 1, $"got {after.Finished.Count}");
            Check("its output paths were captured",
                after.Finished.FirstOrDefault()?.OutputPaths.Count == 1,
                $"got {after.Finished.FirstOrDefault()?.OutputPaths.Count}");
        }
        finally
        {
            try { Directory.Delete(root, recursive: true); } catch (IOException) { }
        }

        return Report();
    }

    private static int Report()
    {
        Console.WriteLine();
        Console.WriteLine(_failures == 0 ? "LANE AGGREGATION OK" : $"LANE AGGREGATION FAILED: {_failures}");
        return _failures == 0 ? 0 : 1;
    }

    private static string Event(string name, string jobId, string preset, string workspace, string[] files) =>
        JsonSerializer.Serialize(new
        {
            ts = DateTimeOffset.UtcNow,
            ev = name,
            jobId,
            preset,
            workspace,
            files,
        });

    private static string Variant(
        string jobId, string preset, string workspace, string file, int n, int total, string output) =>
        JsonSerializer.Serialize(new
        {
            ts = DateTimeOffset.UtcNow,
            ev = "job.variant",
            jobId,
            preset,
            workspace,
            file,
            n,
            total,
            output,
        });

    private static string Done(string jobId, string preset, string workspace, int outputs) =>
        JsonSerializer.Serialize(new
        {
            ts = DateTimeOffset.UtcNow,
            ev = "job.done",
            jobId,
            preset,
            workspace,
            outputs,
        });

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
