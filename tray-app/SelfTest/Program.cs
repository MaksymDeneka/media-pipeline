using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using MediaPipelineTray.Services;

namespace MediaPipelineTray.SelfTest;

/// <summary>
/// Verifies chunked upload without a remote.
///
/// The two parts that would silently corrupt data are the split and the remote reassembly
/// script. Both can be exercised locally: the script is ordinary PowerShell, so it runs against
/// a local folder standing in for the remote. That covers everything except the SFTP transport
/// itself, which is rclone's job rather than ours.
/// </summary>
internal static class Program
{
    private static int _failures;

    private static async Task<int> Main(string[] args)
    {
        // The live check needs a reachable remote, so it is opt-in and never part of the
        // default offline run.
        if (args.Contains("--live"))
        {
            return await LiveUploadTest.RunAsync(sizeMB: 12, chunkSizeMB: 3);
        }

        var root = Path.Combine(Path.GetTempPath(), "mp-upload-selftest");

        if (Directory.Exists(root))
        {
            Directory.Delete(root, recursive: true);
        }

        Directory.CreateDirectory(root);

        try
        {
            await SplitProducesExactChunks(root);
            await SplitIsResumable(root);
            await AssemblyReproducesTheFile(root);
            await AssemblyRejectsACorruptedPart(root);
            await AssemblyRejectsAMissingPart(root);
            await AssemblyLeavesNoTempFileOnFailure(root);
            ConfigEditingPreservesTheFile(root);
            ConfigAddsAndRemoves(root);
            _failures += await LaneAggregationTest.RunAsync();
            _failures += ArchiveTest.Run();
        }
        finally
        {
            try { Directory.Delete(root, recursive: true); } catch (IOException) { }
        }

        Console.WriteLine();

        if (_failures == 0)
        {
            Console.WriteLine("UPLOAD SELF-TEST OK");
            return 0;
        }

        Console.WriteLine($"UPLOAD SELF-TEST FAILED: {_failures} check(s)");
        return 1;
    }

    // --- checks ------------------------------------------------------------

    private static async Task SplitProducesExactChunks(string root)
    {
        Section("Split produces exact chunks");

        var (source, _) = MakeSource(root, "split", bytes: 5 * 1024 * 1024 + 12345);
        var parts = Path.Combine(root, "split-parts");
        Directory.CreateDirectory(parts);

        var chunks = FileChunker.Plan(source, chunkSizeMB: 1);

        Check("plans 6 chunks for 5 MB plus a remainder", chunks.Count == 6, $"got {chunks.Count}");
        Check("last chunk is the remainder", chunks[^1].Length == 12345, $"got {chunks[^1].Length}");
        Check("total equals the source length",
            chunks.Sum(c => c.Length) == new FileInfo(source).Length, "sum mismatch");
        Check("part names are 1-based and zero-padded",
            chunks[0].FileName.EndsWith(".part00001"), chunks[0].FileName);

        foreach (var chunk in chunks)
        {
            await FileChunker.WritePartAsync(source, chunk, Path.Combine(parts, chunk.FileName), 1, default);
        }

        var written = Directory.GetFiles(parts).Length;
        Check("writes one file per chunk", written == chunks.Count, $"got {written}");

        var lengthsMatch = chunks.All(c => new FileInfo(Path.Combine(parts, c.FileName)).Length == c.Length);
        Check("every part is exactly its planned length", lengthsMatch, "a part is the wrong size");
    }

    private static async Task SplitIsResumable(string root)
    {
        Section("Split is resumable");

        var (source, _) = MakeSource(root, "resume", bytes: 3 * 1024 * 1024);
        var parts = Path.Combine(root, "resume-parts");
        Directory.CreateDirectory(parts);

        var chunks = FileChunker.Plan(source, chunkSizeMB: 1);

        foreach (var chunk in chunks)
        {
            await FileChunker.WritePartAsync(source, chunk, Path.Combine(parts, chunk.FileName), 1, default);
        }

        var rewrote = await FileChunker.WritePartAsync(
            source, chunks[0], Path.Combine(parts, chunks[0].FileName), 1, default);

        Check("a complete part is reused rather than rewritten", !rewrote, "it rewrote the part");

        // A truncated part is the interesting case: it must not be trusted.
        var truncated = Path.Combine(parts, chunks[1].FileName);
        await File.WriteAllBytesAsync(truncated, new byte[10]);

        var rewroteTruncated = await FileChunker.WritePartAsync(source, chunks[1], truncated, 1, default);

        Check("a short part is rewritten", rewroteTruncated, "it kept the truncated part");
        Check("the rewritten part is the right length",
            new FileInfo(truncated).Length == chunks[1].Length, "still wrong length");
    }

    private static async Task AssemblyReproducesTheFile(string root)
    {
        Section("Remote assembly reproduces the file byte for byte");

        var result = await AssembleLocally(root, "good", corrupt: false, dropPart: false);

        Check("the assembly script succeeds", result.ExitCode == 0, result.Output);
        Check("the assembled file exists", File.Exists(result.FinalPath), "missing output");

        if (File.Exists(result.FinalPath))
        {
            Check("the assembled file is byte-identical",
                Hash(result.FinalPath) == Hash(result.SourcePath),
                "hash mismatch between source and reassembled file");

            Check("the assembled file is the right length",
                new FileInfo(result.FinalPath).Length == new FileInfo(result.SourcePath).Length,
                "length mismatch");
        }

        // The script consumes the parts but leaves the folder, because it is running from it.
        // Removing the folder is the caller's second, short SSH command.
        var leftovers = Directory.Exists(result.PartsPath)
            ? Directory.GetFiles(result.PartsPath).Select(Path.GetFileName).ToArray()
            : [];

        Check("the consumed parts are deleted",
            leftovers.All(f => f is "manifest.json" or "assemble.ps1"),
            "left behind: " + string.Join(", ", leftovers));
    }

    private static async Task AssemblyRejectsACorruptedPart(string root)
    {
        Section("Remote assembly rejects a corrupted part");

        // The old script compared lengths only, so a same-size corruption slipped through.
        var result = await AssembleLocally(root, "corrupt", corrupt: true, dropPart: false);

        Check("the assembly script fails", result.ExitCode != 0, "it accepted a corrupt part");
        Check("it says which part failed its checksum",
            result.Output.Contains("checksum", StringComparison.OrdinalIgnoreCase),
            result.Output);
        Check("no file is left in place", !File.Exists(result.FinalPath), "it wrote an output anyway");
    }

    private static async Task AssemblyRejectsAMissingPart(string root)
    {
        Section("Remote assembly rejects a missing part");

        var result = await AssembleLocally(root, "missing", corrupt: false, dropPart: true);

        Check("the assembly script fails", result.ExitCode != 0, "it accepted a missing part");
        Check("it names the missing part",
            result.Output.Contains("Missing part", StringComparison.OrdinalIgnoreCase),
            result.Output);
    }

    private static async Task AssemblyLeavesNoTempFileOnFailure(string root)
    {
        Section("A failed assembly leaves no temp file behind");

        var result = await AssembleLocally(root, "notemp", corrupt: true, dropPart: false);

        var temp = result.FinalPath + ".chunked.tmp";
        Check("the .chunked.tmp is removed", !File.Exists(temp),
            "the old script left this behind forever and nothing cleaned it up");
    }


    // --- config editing ----------------------------------------------------

    private const string SampleConfig = """
        ; ===========================================================
        ;  Media Pipeline - Settings
        ; ===========================================================

        [General]
        ; Where all the watcher's folders live.
        PipelineRoot = D:\MediaPipeline

        [preset photos]
        VideoCopies = 0
        ImageCopies = 100

        [Video]
        ; Lower CRF = better quality + bigger files.
        Crf = 24              ; default: 24
        MaxWidth = 1080
        """;

    private static void ConfigEditingPreservesTheFile(string root)
    {
        Section("Config editing preserves the file");

        var path = Path.Combine(root, "config.ini");
        File.WriteAllText(path, SampleConfig);

        var original = File.ReadAllLines(path);

        var ini = IniFile.Load(path);
        ini.Set("Crf", "20");
        ini.Save(path);

        var edited = File.ReadAllLines(path);

        Check("the line count is unchanged", edited.Length == original.Length,
            $"{original.Length} became {edited.Length}");

        Check("the banner comment survives",
            edited.Any(l => l.Contains("Media Pipeline - Settings")), "banner lost");

        Check("the explanatory comment above the key survives",
            edited.Any(l => l.Contains("Lower CRF = better quality")), "comment lost");

        var crfLine = edited.First(l => l.TrimStart().StartsWith("Crf"));
        Check("the value is updated", crfLine.Contains("20"), crfLine);
        Check("the trailing comment survives", crfLine.Contains("; default: 24"), crfLine);

        // Reading back is the real proof: the watcher parses this same shape.
        var reread = IniFile.Load(path);
        var globals = reread.ReadGlobals();

        Check("globals reread correctly", globals["Crf"] == "20", globals.GetValueOrDefault("Crf", "<missing>"));
        Check("an untouched global is intact", globals["MaxWidth"] == "1080",
            globals.GetValueOrDefault("MaxWidth", "<missing>"));
        Check("preset keys stay out of globals", !globals.ContainsKey("ImageCopies"),
            "a preset key leaked into globals");

        var presets = reread.ReadPresets();
        Check("the preset is found", presets.ContainsKey("photos"), "photos missing");
        Check("preset values are read", presets["photos"]["ImageCopies"] == "100",
            presets["photos"].GetValueOrDefault("ImageCopies", "<missing>"));
    }

    private static void ConfigAddsAndRemoves(string root)
    {
        Section("Config adds and removes presets");

        var path = Path.Combine(root, "config-2.ini");
        File.WriteAllText(path, SampleConfig);

        var ini = IniFile.Load(path);
        ini.AddPreset("clips");
        ini.Set("VideoCopies", "4", "clips");
        ini.Set("ImageCopies", "0", "clips");
        ini.Save(path);

        var afterAdd = IniFile.Load(path).ReadPresets();
        Check("the new preset exists", afterAdd.ContainsKey("clips"), "clips missing");
        Check("its values are written", afterAdd["clips"]["VideoCopies"] == "4",
            afterAdd["clips"].GetValueOrDefault("VideoCopies", "<missing>"));
        Check("the existing preset is untouched", afterAdd.ContainsKey("photos"), "photos lost");

        // Setting a preset key must not touch the global of the same name.
        var globalsAfterAdd = IniFile.Load(path).ReadGlobals();
        Check("a preset edit does not create a global",
            !globalsAfterAdd.ContainsKey("VideoCopies"), "VideoCopies leaked to globals");

        var ini2 = IniFile.Load(path);
        ini2.RemoveKey("ImageCopies", "clips");
        ini2.Save(path);

        var afterRemoveKey = IniFile.Load(path).ReadPresets();
        Check("removing a key drops just that key",
            !afterRemoveKey["clips"].ContainsKey("ImageCopies"), "key remains");
        Check("its sibling survives", afterRemoveKey["clips"]["VideoCopies"] == "4", "sibling lost");

        var ini3 = IniFile.Load(path);
        ini3.RemovePreset("clips");
        ini3.Save(path);

        var afterRemove = IniFile.Load(path);
        Check("removing a preset drops the section",
            !afterRemove.ReadPresets().ContainsKey("clips"), "clips remains");
        Check("other presets survive", afterRemove.ReadPresets().ContainsKey("photos"), "photos lost");
        Check("globals survive a section removal",
            afterRemove.ReadGlobals()["MaxWidth"] == "1080", "globals damaged");
    }

    // --- helpers -----------------------------------------------------------


    private sealed record AssemblyResult(
        int ExitCode, string Output, string SourcePath, string FinalPath, string PartsPath);

    /// <summary>
    /// Runs the real remote assembly script against local folders. The script only ever touches
    /// paths from its manifest, so pointing those at a temp folder exercises it faithfully.
    /// </summary>
    private static async Task<AssemblyResult> AssembleLocally(
        string root, string name, bool corrupt, bool dropPart)
    {
        var (source, fileName) = MakeSource(root, name, bytes: 2 * 1024 * 1024 + 999);

        var partsPath = Path.Combine(root, name + "-parts");
        var destination = Path.Combine(root, name + "-dest");
        Directory.CreateDirectory(partsPath);
        Directory.CreateDirectory(destination);

        var chunks = FileChunker.Plan(source, chunkSizeMB: 1).ToList();

        foreach (var chunk in chunks)
        {
            var partPath = Path.Combine(partsPath, chunk.FileName);
            await FileChunker.WritePartAsync(source, chunk, partPath, 1, default);
            chunk.Sha256 = await FileChunker.HashAsync(partPath, default);
        }

        if (corrupt)
        {
            // Flip a byte without changing the length, which length checks cannot catch.
            var target = Path.Combine(partsPath, chunks[1].FileName);
            var bytes = await File.ReadAllBytesAsync(target);
            bytes[0] ^= 0xFF;
            await File.WriteAllBytesAsync(target, bytes);
        }

        if (dropPart)
        {
            File.Delete(Path.Combine(partsPath, chunks[^1].FileName));
        }

        var manifest = new
        {
            fileName,
            expectedLength = new FileInfo(source).Length,
            chunkCount = chunks.Count,
            remoteDirectory = destination,
            remotePartsDirectory = partsPath,
            parts = chunks.Select(c => new { name = c.FileName, length = c.Length, sha256 = c.Sha256 }),
        };

        // Same contract as a real upload: manifest.json sits beside the script, and the
        // script runs from the parts folder.
        await File.WriteAllTextAsync(
            Path.Combine(partsPath, "manifest.json"), JsonSerializer.Serialize(manifest));

        var scriptPath = Path.Combine(partsPath, "assemble.ps1");
        await File.WriteAllTextAsync(scriptPath, UploadService.BuildRemoteScript());

        var (exitCode, output) = await RunPowerShell(scriptPath);

        return new AssemblyResult(
            exitCode, output, source, Path.Combine(destination, fileName), partsPath);
    }

    private static (string Path, string Name) MakeSource(string root, string name, int bytes)
    {
        var fileName = name + ".bin";
        var path = Path.Combine(root, fileName);

        // Deterministic pseudo-random content: compresses badly, and reproduces on rerun.
        var data = new byte[bytes];
        var random = new Random(Seed: bytes);
        random.NextBytes(data);
        File.WriteAllBytes(path, data);

        return (path, fileName);
    }

    private static string Hash(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static async Task<(int ExitCode, string Output)> RunPowerShell(string scriptPath)
    {
        var candidates = new[]
        {
            @"C:\Tools\pwsh\pwsh.exe",
            "pwsh.exe",
            @"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
        };

        var executable = candidates.FirstOrDefault(File.Exists) ?? candidates[^1];

        var info = new ProcessStartInfo
        {
            FileName = executable,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };

        info.ArgumentList.Add("-NoProfile");
        info.ArgumentList.Add("-File");
        info.ArgumentList.Add(scriptPath);

        using var process = Process.Start(info)!;

        var stdout = await process.StandardOutput.ReadToEndAsync();
        var stderr = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();

        return (process.ExitCode, stdout + stderr);
    }

    private static void Section(string title) => Console.WriteLine($"\n== {title}");

    private static void Check(string label, bool passed, string detail)
    {
        if (passed)
        {
            Console.WriteLine($"   PASS  {label}");
            return;
        }

        _failures++;
        Console.WriteLine($"   FAIL  {label}");

        var trimmed = detail.Trim();
        if (trimmed.Length > 0)
        {
            Console.WriteLine($"         {trimmed.Split('\n')[^1].Trim()}");
        }
    }
}
