using System.Diagnostics;
using System.Security.Cryptography;
using MediaPipelineTray.Services;

namespace MediaPipelineTray.SelfTest;

/// <summary>
/// A real round-trip against the configured remote.
///
/// Everything else in this project runs offline. This is the one check that exercises the
/// actual SFTP transport and the actual SSH assembly, because those are the parts no local
/// stand-in can prove.
///
/// It uploads its own small file into a scratch folder, verifies the reassembled bytes match,
/// and deletes what it made. It never touches the real sync folder, and it never uploads
/// anything it did not create.
/// </summary>
internal static class LiveUploadTest
{
    private const string ScratchFolder = "_uploadselftest";

    public static async Task<int> RunAsync(int sizeMB, int chunkSizeMB)
    {
        Console.WriteLine("== Live upload round-trip");
        Console.WriteLine($"   {sizeMB} MB file in {chunkSizeMB} MB chunks");

        var paths = PipelinePaths.Discover();
        var globals = IniFile.Load(paths.ConfigFile).ReadGlobals();
        var configured = UploadTarget.FromConfig(globals);

        // Everything lives under one scratch folder so the real sync folder is untouched and
        // cleanup is a single delete. All three roots must move together: the SFTP root is
        // where parts are written, the Windows parts root is where the remote script is run
        // from, and they have to be the same place.
        var scratchWindows = configured.RemotePartsRoot.TrimEnd('\\') + "\\" + ScratchFolder;
        var scratchSftp = configured.RemoteSftpPartsRoot.TrimEnd('/') + "/" + ScratchFolder;

        var target = configured with
        {
            ChunkSizeMB = chunkSizeMB,
            RemotePartsRoot = scratchWindows + "\\parts",
            RemoteSftpPartsRoot = scratchSftp + "/parts",
            RemoteDirectory = scratchWindows + "\\out",

            // Exercised here rather than trusted: this deletes the user's original, so it must
            // only happen once the remote copy has been read back and confirmed.
            DeleteAfterUpload = true,
        };

        // Cleanup removes the whole scratch folder, not just the assembled output.
        var scratchRoot = scratchWindows;

        Console.WriteLine($"   remote      {target.RemoteName}");
        Console.WriteLine($"   assembles   {target.RemoteDirectory}");
        Console.WriteLine($"   ssh         {target.SshHost}:{target.SshPort}");
        Console.WriteLine();

        var localFile = Path.Combine(Path.GetTempPath(), $"upload-selftest-{Guid.NewGuid():n}.bin");
        var failures = 0;

        try
        {
            var data = new byte[sizeMB * 1024 * 1024];
            new Random(Seed: sizeMB).NextBytes(data);
            await File.WriteAllBytesAsync(localFile, data);

            var expectedHash = Convert.ToHexString(SHA256.HashData(data));
            Console.WriteLine($"   local sha   {expectedHash[..16]}...");

            var service = new UploadService(paths);
            var job = new UploadJob { SourcePath = localFile, Target = target };

            var lastPhase = UploadPhase.Queued;
            service.Progress += (_, j) =>
            {
                if (j.Phase != lastPhase)
                {
                    lastPhase = j.Phase;
                    Console.WriteLine($"   {j.Phase.ToString().ToLowerInvariant()} ...");
                }
            };

            var stopwatch = Stopwatch.StartNew();
            await service.RunAsync(job, CancellationToken.None);
            stopwatch.Stop();

            failures += Check("upload completes", job.Phase == UploadPhase.Done, job.Error ?? "");
            failures += Check("every chunk sent",
                job.ChunksSent == job.Chunks.Count,
                $"{job.ChunksSent} of {job.Chunks.Count}");

            if (job.Phase == UploadPhase.Done)
            {
                Console.WriteLine($"   {job.Chunks.Count} chunk(s) in {stopwatch.Elapsed.TotalSeconds:0.0}s");

                // Uploads land in a workspace folder inside the remote directory, so ask the
                // job where it actually put the file rather than assuming.
                var remoteFile = Path.Combine(job.RemoteWorkspaceDirectory, Path.GetFileName(localFile));
                var remoteHash = await RemoteHash(target, remoteFile);

                failures += Check("remote file matches byte for byte",
                    string.Equals(remoteHash, expectedHash, StringComparison.OrdinalIgnoreCase),
                    $"remote {Short(remoteHash)} vs local {Short(expectedHash)}");

                failures += Check("local parts cleaned up",
                    !Directory.Exists(Path.Combine(paths.PipelineRoot, ".sync-parts",
                        Path.GetFileName(localFile) + ".parts")),
                    "local parts remain");

                failures += Check("remote size was verified", job.RemoteVerified, "not verified");

                failures += Check("local original deleted after verification",
                    job.SourceDeleted && !File.Exists(localFile),
                    job.SourceDeleted ? "flag set but file remains" : "not deleted");
            }
        }
        finally
        {
            // The upload deletes this itself when DeleteAfterUpload is on; only clean up
            // after a failure that left it behind.
            if (File.Exists(localFile))
            {
                File.Delete(localFile);
            }

            // Always clean the remote scratch area, even after a failure.
            await CleanRemote(target, scratchRoot);
            Console.WriteLine("   remote scratch folder removed");
        }

        Console.WriteLine();

        if (failures == 0)
        {
            Console.WriteLine("LIVE UPLOAD OK");
            return 0;
        }

        Console.WriteLine($"LIVE UPLOAD FAILED: {failures} check(s)");
        return 1;
    }

    private static string Short(string hash) => hash.Length >= 16 ? hash[..16] + "..." : hash;

    /// <summary>
    /// Runs PowerShell on the remote via -EncodedCommand.
    ///
    /// An SSH command reaches a Windows host through cmd.exe, which mangles nested quoting.
    /// Base64 sidesteps that, and these commands are far short of the 8191 character limit.
    /// </summary>
    private static async Task<string> RemotePowerShell(UploadTarget target, string script)
    {
        var encoded = Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(script));
        var (_, output) = await Ssh(target, $"powershell -NoProfile -EncodedCommand {encoded}");
        return output;
    }

    private static async Task<string> RemoteHash(UploadTarget target, string remotePath)
    {
        var output = await RemotePowerShell(
            target, $"(Get-FileHash -LiteralPath '{remotePath}' -Algorithm SHA256).Hash");

        // Remote PowerShell wraps its output in a CLIXML progress blob, so find the line that
        // actually looks like a SHA-256 rather than trusting its position.
        return output
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .FirstOrDefault(line => line.Length == 64 && line.All(Uri.IsHexDigit))
            ?? "<no hash found>";
    }

    private static async Task CleanRemote(UploadTarget target, string scratchRoot) =>
        await RemotePowerShell(
            target,
            $"Remove-Item -LiteralPath '{scratchRoot}' -Recurse -Force -ErrorAction SilentlyContinue");

    private static async Task<(int ExitCode, string Output)> Ssh(UploadTarget target, string command)
    {
        var info = new ProcessStartInfo
        {
            FileName = "ssh",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };

        foreach (var argument in new[]
        {
            "-o", "BatchMode=yes",
            "-o", "ConnectTimeout=10",
            "-i", target.SshKeyFile,
            "-p", target.SshPort.ToString(),
            target.SshHost,
            command,
        })
        {
            info.ArgumentList.Add(argument);
        }

        using var process = Process.Start(info)!;
        var stdout = await process.StandardOutput.ReadToEndAsync();
        var stderr = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();

        return (process.ExitCode, stdout + stderr);
    }

    private static int Check(string label, bool passed, string detail)
    {
        if (passed)
        {
            Console.WriteLine($"   PASS  {label}");
            return 0;
        }

        Console.WriteLine($"   FAIL  {label}");

        var trimmed = detail.Trim();
        if (trimmed.Length > 0)
        {
            Console.WriteLine($"         {trimmed.Split('\n')[^1].Trim()}");
        }

        return 1;
    }
}
