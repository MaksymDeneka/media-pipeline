using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace MediaPipelineTray.Services;

public enum ChunkState
{
    Pending,
    Splitting,
    Sending,
    Sent,
    Failed,
}

public enum UploadPhase
{
    Queued,
    Splitting,
    Sending,
    Assembling,
    Verifying,
    Done,
    Failed,
    Cancelled,
    Paused,
}

public sealed class ChunkProgress
{
    public required int Index { get; init; }
    public required long Length { get; init; }
    public required string FileName { get; init; }
    public string Sha256 { get; set; } = "";
    public ChunkState State { get; set; } = ChunkState.Pending;
    public int Attempts { get; set; }
    public string? Error { get; set; }
}

/// <summary>Where a chunked upload is going, and how it should get there.</summary>
public sealed record UploadTarget
{
    public string RemoteName { get; init; } = "heatup-remote";
    public string RemoteSftpPartsRoot { get; init; } = "/D:/MediaPipeline/.sync-parts";
    public string RemotePartsRoot { get; init; } = @"D:\MediaPipeline\.sync-parts";
    public string RemoteDirectory { get; init; } = @"D:\MediaPipeline\sync";
    public string SshHost { get; init; } = "heatup-remote";
    public int SshPort { get; init; } = 2222;
    public string SshKeyFile { get; init; } = "";
    public int ChunkSizeMB { get; init; } = 256;
    public int ParallelChunks { get; init; } = 4;

    public static UploadTarget FromConfig(IReadOnlyDictionary<string, string> globals)
    {
        var fallbackKey = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".ssh", "heatup_remote_debug_ed25519");

        return new UploadTarget
        {
            RemoteName = Get("RemoteName", "heatup-remote"),
            RemoteSftpPartsRoot = Get("RemoteSftpPartsRoot", "/D:/MediaPipeline/.sync-parts"),
            RemotePartsRoot = Get("RemotePartsRoot", @"D:\MediaPipeline\.sync-parts"),
            RemoteDirectory = Get("RemoteDirectory", @"D:\MediaPipeline\sync"),
            SshHost = Get("RemoteSshHost", "heatup-remote"),
            SshPort = int.TryParse(Get("RemoteSshPort", "2222"), out var port) ? port : 2222,
            SshKeyFile = Environment.ExpandEnvironmentVariables(Get("RemoteSshKeyFile", fallbackKey)),
            ChunkSizeMB = int.TryParse(Get("ChunkSizeMB", "256"), out var size) ? size : 256,
            ParallelChunks = int.TryParse(Get("ParallelChunks", "4"), out var parallel) ? parallel : 4,
        };

        string Get(string key, string fallback) =>
            globals.TryGetValue(key, out var value) && value.Length > 0 ? value : fallback;
    }
}

/// <summary>One file being chunked and uploaded.</summary>
public sealed class UploadJob
{
    public required string SourcePath { get; init; }
    public required UploadTarget Target { get; init; }

    public string FileName => Path.GetFileName(SourcePath);

    /// <summary>
    /// Taken from the staging folder: a file in sync\LC belongs to LC. The remote is split by
    /// workspace and not by preset, so this is the only routing an upload needs.
    /// </summary>
    public string Workspace => Path.GetFileName(Path.GetDirectoryName(SourcePath) ?? "") is { Length: > 0 } name
        ? name
        : "general";

    /// <summary>Where this file lands on the remote, inside its workspace folder.</summary>
    public string RemoteWorkspaceDirectory =>
        Path.Combine(Target.RemoteDirectory, Workspace);
    public long TotalBytes { get; set; }
    public List<ChunkProgress> Chunks { get; } = [];
    public UploadPhase Phase { get; set; } = UploadPhase.Queued;
    public string? Error { get; set; }
    public DateTimeOffset? StartedUtc { get; set; }

    public int ChunksSent => Chunks.Count(c => c.State == ChunkState.Sent);
    public long BytesSent => Chunks.Where(c => c.State == ChunkState.Sent).Sum(c => c.Length);

    public double Fraction => TotalBytes > 0 ? (double)BytesSent / TotalBytes : 0;
}

/// <summary>
/// Chunked upload over SFTP, reimplemented from sync-upload-chunked-file.ps1.
///
/// The shape is the same because the constraint is the same: the link drops on large single
/// transfers, so the file is split, the parts are sent independently, and the remote side
/// reassembles them.
///
/// What is different, and why:
///
///  - Parts are sent one rclone call at a time rather than one call for the whole folder, so
///    a failure is attributable to a chunk and only that chunk is retried. The old script
///    failed the entire run.
///  - Every part carries a SHA-256 that the remote verifies before appending. The old script
///    compared lengths only, with hash checking explicitly disabled, so a same-size-but-corrupt
///    part would have been assembled in silently.
///  - A failed assembly cleans up its own .chunked.tmp. The old script left it on the remote
///    forever, and nothing ever removed it.
/// </summary>
public sealed class UploadService
{
    private const string ManifestFileName = "manifest.json";
    private const string AssemblyScriptName = "assemble.ps1";

    private readonly PipelinePaths _paths;

    public UploadService(PipelinePaths paths) => _paths = paths;

    public event EventHandler<UploadJob>? Progress;

    private void Report(UploadJob job) => Progress?.Invoke(this, job);

    private string LocalPartsDirectory(UploadJob job) =>
        Path.Combine(_paths.PipelineRoot, ".sync-parts", job.FileName + ".parts");

    public async Task RunAsync(UploadJob job, CancellationToken cancellationToken)
    {
        job.StartedUtc = DateTimeOffset.UtcNow;

        try
        {
            await SplitAsync(job, cancellationToken).ConfigureAwait(false);
            await SendAsync(job, cancellationToken).ConfigureAwait(false);
            await AssembleAsync(job, cancellationToken).ConfigureAwait(false);

            Cleanup(job);

            job.Phase = UploadPhase.Done;
        }
        catch (OperationCanceledException)
        {
            // Local parts are deliberately kept so the next run resumes instead of restarting.
            job.Phase = UploadPhase.Cancelled;
        }
        catch (Exception ex)
        {
            job.Phase = UploadPhase.Failed;
            job.Error = ex.Message;
        }

        Report(job);
    }

    // --- split -------------------------------------------------------------

    private async Task SplitAsync(UploadJob job, CancellationToken cancellationToken)
    {
        job.Phase = UploadPhase.Splitting;
        Report(job);

        job.TotalBytes = new FileInfo(job.SourcePath).Length;

        var partsDirectory = LocalPartsDirectory(job);
        Directory.CreateDirectory(partsDirectory);

        job.Chunks.Clear();
        job.Chunks.AddRange(FileChunker.Plan(job.SourcePath, job.Target.ChunkSizeMB));
        Report(job);

        foreach (var chunk in job.Chunks)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var partPath = Path.Combine(partsDirectory, chunk.FileName);

            chunk.State = ChunkState.Splitting;
            Report(job);

            await FileChunker
                .WritePartAsync(job.SourcePath, chunk, partPath, job.Target.ChunkSizeMB, cancellationToken)
                .ConfigureAwait(false);

            chunk.Sha256 = await FileChunker.HashAsync(partPath, cancellationToken).ConfigureAwait(false);
            chunk.State = ChunkState.Pending;
            Report(job);
        }

        await WriteAssemblyFilesAsync(job, partsDirectory, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Writes the manifest and the assembly script into the parts folder so they travel with
    /// the parts.
    ///
    /// They are files rather than command-line arguments because the remote runs Windows: an
    /// SSH command goes through cmd.exe, which caps at 8191 characters. A manifest carrying a
    /// SHA-256 per chunk blows that at four chunks, and a real upload has hundreds.
    /// </summary>
    private static async Task WriteAssemblyFilesAsync(
        UploadJob job, string partsDirectory, CancellationToken cancellationToken)
    {
        var manifest = new
        {
            fileName = job.FileName,
            expectedLength = job.TotalBytes,
            chunkCount = job.Chunks.Count,
            remoteDirectory = job.RemoteWorkspaceDirectory,
            remotePartsDirectory = Path.Combine(job.Target.RemotePartsRoot, job.FileName + ".parts"),
            parts = job.Chunks.Select(c => new { name = c.FileName, length = c.Length, sha256 = c.Sha256 }),
        };

        await File.WriteAllTextAsync(
            Path.Combine(partsDirectory, ManifestFileName),
            JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }),
            cancellationToken).ConfigureAwait(false);

        await File.WriteAllTextAsync(
            Path.Combine(partsDirectory, AssemblyScriptName),
            BuildRemoteScript(),
            cancellationToken).ConfigureAwait(false);
    }

    // --- send --------------------------------------------------------------

    private async Task SendAsync(UploadJob job, CancellationToken cancellationToken)
    {
        job.Phase = UploadPhase.Sending;
        Report(job);

        var partsDirectory = LocalPartsDirectory(job);
        var remoteParts = $"{job.Target.RemoteName}:{job.Target.RemoteSftpPartsRoot.TrimEnd('/')}/{job.FileName}.parts";

        await RunProcessAsync("rclone", ["mkdir", remoteParts, "--timeout", "30s"], cancellationToken)
            .ConfigureAwait(false);

        using var gate = new SemaphoreSlim(Math.Max(1, job.Target.ParallelChunks));

        var sends = job.Chunks
            .Where(chunk => chunk.State != ChunkState.Sent)
            .Select(async chunk =>
            {
                await gate.WaitAsync(cancellationToken).ConfigureAwait(false);

                try
                {
                    await SendChunkAsync(job, chunk, partsDirectory, remoteParts, cancellationToken)
                        .ConfigureAwait(false);
                }
                finally
                {
                    gate.Release();
                }
            });

        await Task.WhenAll(sends).ConfigureAwait(false);

        var failed = job.Chunks.Where(chunk => chunk.State == ChunkState.Failed).ToList();
        if (failed.Count > 0)
        {
            throw new InvalidOperationException(
                $"{failed.Count} chunk(s) could not be sent. First error: {failed[0].Error}");
        }

        // The manifest and the script go last, so their presence means the parts are complete.
        foreach (var name in new[] { ManifestFileName, AssemblyScriptName })
        {
            await RunProcessAsync(
                "rclone",
                [
                    "copyto", Path.Combine(partsDirectory, name), $"{remoteParts}/{name}",
                    "--retries", "3", "--timeout", "5m", "--contimeout", "30s",
                    "--sftp-disable-hashcheck",
                ],
                cancellationToken).ConfigureAwait(false);
        }
    }

    private async Task SendChunkAsync(
        UploadJob job,
        ChunkProgress chunk,
        string partsDirectory,
        string remoteParts,
        CancellationToken cancellationToken)
    {
        const int maxAttempts = 5;
        var localPath = Path.Combine(partsDirectory, chunk.FileName);

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            chunk.Attempts = attempt;
            chunk.State = ChunkState.Sending;
            Report(job);

            var result = await RunProcessAsync(
                "rclone",
                [
                    "copyto", localPath, $"{remoteParts}/{chunk.FileName}",
                    "--retries", "2",
                    "--low-level-retries", "10",
                    "--timeout", "10m",
                    "--contimeout", "30s",
                    "--sftp-disable-hashcheck",
                ],
                cancellationToken,
                throwOnFailure: false).ConfigureAwait(false);

            if (result.ExitCode == 0)
            {
                chunk.State = ChunkState.Sent;
                chunk.Error = null;
                Report(job);
                return;
            }

            chunk.Error = result.Output.Length > 0
                ? result.Output.Split('\n')[^1].Trim()
                : $"rclone exited with {result.ExitCode}";

            if (attempt < maxAttempts)
            {
                // Back off a little; a dropped link usually needs a moment, not an instant retry.
                await Task.Delay(TimeSpan.FromSeconds(2 * attempt), cancellationToken).ConfigureAwait(false);
            }
        }

        chunk.State = ChunkState.Failed;
        Report(job);
    }

    // --- assemble ----------------------------------------------------------

    private async Task AssembleAsync(UploadJob job, CancellationToken cancellationToken)
    {
        job.Phase = UploadPhase.Assembling;
        Report(job);

        var remotePartsWindows = Path.Combine(job.Target.RemotePartsRoot, job.FileName + ".parts");
        var scriptPath = Path.Combine(remotePartsWindows, AssemblyScriptName);

        // A fixed, short command regardless of how many chunks there are.
        var result = await SshAsync(
            job.Target,
            $"powershell -NoProfile -ExecutionPolicy Bypass -File \"{scriptPath}\"",
            cancellationToken).ConfigureAwait(false);

        if (result.ExitCode != 0)
        {
            throw new InvalidOperationException($"Remote assembly failed: {result.Output.Trim()}");
        }

        // The script removes the parts it consumed but not its own folder, which it is running
        // from. Clearing that is a second, equally short command.
        await SshAsync(
            job.Target,
            $"powershell -NoProfile -Command \"Remove-Item -LiteralPath '{remotePartsWindows}' -Recurse -Force -ErrorAction SilentlyContinue\"",
            cancellationToken).ConfigureAwait(false);
    }

    private static async Task<ProcessResult> SshAsync(
        UploadTarget target, string command, CancellationToken cancellationToken) =>
        await RunProcessAsync(
            "ssh",
            [
                "-o", "BatchMode=yes",
                "-o", "ConnectTimeout=8",
                "-o", "ServerAliveInterval=30",
                "-o", "ServerAliveCountMax=3",
                "-o", "TCPKeepAlive=yes",
                "-i", target.SshKeyFile,
                "-p", target.SshPort.ToString(),
                target.SshHost,
                command,
            ],
            cancellationToken,
            throwOnFailure: false).ConfigureAwait(false);

    /// <summary>
    /// The reassembly script that runs on the remote host.
    ///
    /// It reads manifest.json from its own folder rather than taking it as an argument,
    /// because the remote runs Windows and an SSH command goes through cmd.exe, which caps at
    /// 8191 characters. With a SHA-256 per chunk an inline manifest exceeds that at four
    /// chunks, and a real upload has hundreds.
    ///
    /// Every part is hash-checked before it is appended, and the temporary file is removed on
    /// any failure, so a failed run leaves nothing behind to clean up by hand.
    /// </summary>
    public static string BuildRemoteScript() =>
        """
        $ErrorActionPreference = 'Stop'

        $manifestPath = Join-Path $PSScriptRoot 'manifest.json'
        if (-not (Test-Path -LiteralPath $manifestPath)) {
            throw "No manifest beside the script at $manifestPath"
        }

        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

        $partsDirectory = $PSScriptRoot
        $finalPath = Join-Path $manifest.remoteDirectory $manifest.fileName
        $tempPath = "$finalPath.chunked.tmp"

        if (-not (Test-Path -LiteralPath $manifest.remoteDirectory)) {
            New-Item -ItemType Directory -Path $manifest.remoteDirectory -Force | Out-Null
        }

        try {
            $stream = [IO.File]::Create($tempPath)

            try {
                foreach ($part in $manifest.parts) {
                    $partPath = Join-Path $partsDirectory $part.name

                    if (-not (Test-Path -LiteralPath $partPath)) {
                        throw "Missing part: $($part.name)"
                    }

                    $actualLength = (Get-Item -LiteralPath $partPath).Length
                    if ($actualLength -ne $part.length) {
                        throw "Part $($part.name) is $actualLength bytes, expected $($part.length)"
                    }

                    $actualHash = (Get-FileHash -LiteralPath $partPath -Algorithm SHA256).Hash
                    if ($actualHash -ne $part.sha256) {
                        throw "Part $($part.name) failed its checksum"
                    }

                    $bytes = [IO.File]::ReadAllBytes($partPath)
                    $stream.Write($bytes, 0, $bytes.Length)
                }
            }
            finally {
                $stream.Dispose()
            }

            $assembledLength = (Get-Item -LiteralPath $tempPath).Length
            if ($assembledLength -ne $manifest.expectedLength) {
                throw "Assembled $assembledLength bytes, expected $($manifest.expectedLength)"
            }

            Move-Item -LiteralPath $tempPath -Destination $finalPath -Force
            Write-Output "assembled $finalPath ($assembledLength bytes)"

            # Remove the parts, but not this folder: the script is running from it. The caller
            # clears the folder afterwards.
            foreach ($part in $manifest.parts) {
                Remove-Item -LiteralPath (Join-Path $partsDirectory $part.name) -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            # Never leave a half-written file behind.
            if (Test-Path -LiteralPath $tempPath) {
                Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
            }

            throw
        }
        """;

    private void Cleanup(UploadJob job)
    {
        var partsDirectory = LocalPartsDirectory(job);

        try
        {
            if (Directory.Exists(partsDirectory))
            {
                Directory.Delete(partsDirectory, recursive: true);
            }
        }
        catch (IOException)
        {
            // Leaving local parts behind wastes disk but breaks nothing, and retention sweeps
            // .sync-parts anyway.
        }
    }

    // --- process plumbing --------------------------------------------------

    private sealed record ProcessResult(int ExitCode, string Output);

    private static async Task<ProcessResult> RunProcessAsync(
        string fileName,
        string[] arguments,
        CancellationToken cancellationToken,
        bool throwOnFailure = true)
    {
        var info = new ProcessStartInfo
        {
            FileName = fileName,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };

        foreach (var argument in arguments)
        {
            info.ArgumentList.Add(argument);
        }

        using var process = new Process { StartInfo = info };

        // stdout and stderr are delivered on separate threads, and StringBuilder is not
        // thread-safe, so both handlers append under one lock.
        var output = new StringBuilder();
        var outputLock = new object();

        void Append(string? line)
        {
            if (line is null)
            {
                return;
            }

            lock (outputLock)
            {
                output.AppendLine(line);
            }
        }

        process.OutputDataReceived += (_, e) => Append(e.Data);
        process.ErrorDataReceived += (_, e) => Append(e.Data);

        try
        {
            process.Start();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Could not run {fileName}. Is it installed and on PATH? ({ex.Message})", ex);
        }

        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        try
        {
            await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            // Killing the tree matters: rclone spawns children that would otherwise keep the
            // transfer running after a cancel.
            try { process.Kill(entireProcessTree: true); } catch (InvalidOperationException) { }
            throw;
        }

        string captured;
        lock (outputLock)
        {
            captured = output.ToString();
        }

        var result = new ProcessResult(process.ExitCode, captured);

        if (throwOnFailure && result.ExitCode != 0)
        {
            throw new InvalidOperationException($"{fileName} exited with {result.ExitCode}: {result.Output.Trim()}");
        }

        return result;
    }
}
