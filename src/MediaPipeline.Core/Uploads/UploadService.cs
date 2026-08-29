using System.Text;
using System.Text.Json;
using System.Security.Cryptography;
using MediaPipeline.Core.IO;
using MediaPipeline.Core.Tools;

namespace MediaPipeline.Core.Uploads;

/// <summary>
/// Sends files as independently retried, checksummed SFTP chunks and asks the Windows
/// destination to assemble them. Local parts remain after cancellation or failure so the
/// next attempt can reuse them.
/// </summary>
public sealed class UploadService(PipelinePaths paths)
{
    private const string ManifestFileName = "manifest.json";
    private const string UploadLockSuffix = ".upload.lock";
    private static string Rclone => Toolchain.FindRequired("rclone");
    private static string Ssh => Toolchain.FindRequired("ssh");

    public event EventHandler<UploadJob>? Progress;

    public async Task RunAsync(UploadJob job, CancellationToken cancellationToken = default)
    {
        job.StartedUtc = DateTimeOffset.UtcNow;
        job.Error = null;
        ValidateJob(job);

        try
        {
            job.TotalBytes = new FileInfo(job.SourcePath).Length;
            job.SourceSha256 = await FileChunker.HashAsync(job.SourcePath, cancellationToken);
            await using (var uploadLease = AcquireUploadLease(job))
            {
                await SplitAsync(job, cancellationToken);
                await SendAsync(job, cancellationToken);
                await AssembleAsync(job, cancellationToken);
                await VerifyAsync(job, cancellationToken);
                await EnsureSourceUnchangedAsync(job, cancellationToken);
                if (job.Target.DeleteAfterUpload)
                {
                    await ClaimSourceForDeletionAsync(job, cancellationToken);
                }
                Cleanup(job);
            }

            if (job.Target.DeleteAfterUpload)
            {
                DeleteClaimedSource(job);
            }

            job.Phase = UploadPhase.Done;
        }
        catch (OperationCanceledException)
        {
            RestoreClaimedSource(job);
            job.Phase = UploadPhase.Cancelled;
        }
        catch (Exception exception)
        {
            RestoreClaimedSource(job);
            job.Phase = UploadPhase.Failed;
            job.Error = exception.Message;
        }

        Report(job);
    }

    private void Report(UploadJob job) => Progress?.Invoke(this, job);

    private string LocalPartsDirectory(UploadJob job) =>
        Path.Combine(paths.SyncParts, job.Workspace, job.TransferDirectoryName);

    private string LocalLeasePath(UploadJob job) =>
        LocalPartsDirectory(job) + UploadLockSuffix;

    private static void ValidateJob(UploadJob job)
    {
        if (!File.Exists(job.SourcePath))
        {
            throw new FileNotFoundException($"Upload source not found: {job.SourcePath}");
        }

        PipelinePaths.ValidateSegment(job.Workspace, nameof(job.Workspace));
        if (job.FileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
            job.FileName.IndexOfAny(['<', '>', ':', '"', '/', '\\', '|', '?', '*']) >= 0)
        {
            throw new InvalidOperationException($"Invalid upload file name '{job.FileName}'.");
        }
    }

    private async Task SplitAsync(UploadJob job, CancellationToken cancellationToken)
    {
        job.Phase = UploadPhase.Splitting;
        Report(job);

        var partsDirectory = LocalPartsDirectory(job);
        Directory.CreateDirectory(partsDirectory);
        job.Chunks.Clear();
        job.Chunks.AddRange(FileChunker.Plan(job.SourcePath, job.Target.ChunkSizeMB));
        Report(job);

        foreach (var chunk in job.Chunks)
        {
            cancellationToken.ThrowIfCancellationRequested();
            chunk.State = ChunkState.Splitting;
            Report(job);

            var partPath = Path.Combine(partsDirectory, chunk.FileName);
            await FileChunker.WritePartAsync(
                job.SourcePath,
                chunk,
                partPath,
                job.Target.ChunkSizeMB,
                cancellationToken);
            chunk.Sha256 = await FileChunker.HashAsync(partPath, cancellationToken);
            chunk.State = ChunkState.Pending;
            Report(job);
        }

        var partsHash = await FileChunker.HashPartsAsync(
            job.Chunks.Select(chunk => Path.Combine(partsDirectory, chunk.FileName)),
            cancellationToken);
        if (!partsHash.Equals(job.SourceSha256, StringComparison.OrdinalIgnoreCase))
        {
            throw new IOException("The upload source changed while its resumable parts were prepared.");
        }

        await WriteAssemblyFilesAsync(job, partsDirectory, cancellationToken);
    }

    private static async Task WriteAssemblyFilesAsync(
        UploadJob job,
        string partsDirectory,
        CancellationToken cancellationToken)
    {
        var remotePartsDirectory = CombineRemoteWindows(
            job.Target.RemotePartsRoot,
            job.Workspace,
            job.TransferDirectoryName);
        var manifest = new
        {
            fileName = job.FileName,
            expectedLength = job.TotalBytes,
            sourceSha256 = job.SourceSha256,
            chunkCount = job.Chunks.Count,
            remoteDirectory = job.RemoteWorkspaceDirectory,
            remotePartsDirectory,
            destinationLock = DestinationLock(job),
            parts = job.Chunks.Select(chunk => new
            {
                name = chunk.FileName,
                length = chunk.Length,
                sha256 = chunk.Sha256,
            }),
        };

        var manifestPath = Path.Combine(partsDirectory, ManifestFileName);
        await File.WriteAllTextAsync(
            manifestPath,
            JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }),
            cancellationToken);
        job.ManifestSha256 = await FileChunker.HashAsync(manifestPath, cancellationToken);
    }

    private async Task SendAsync(UploadJob job, CancellationToken cancellationToken)
    {
        job.Phase = UploadPhase.Sending;
        Report(job);

        var partsDirectory = LocalPartsDirectory(job);
        var remoteParts =
            $"{job.Target.RemoteName}:{job.Target.RemoteSftpPartsRoot.TrimEnd('/')}/" +
            $"{job.Workspace}/{job.TransferDirectoryName}";
        await RunRequiredAsync(
            Rclone, ["mkdir", remoteParts, "--timeout", "30s"], cancellationToken);

        using var gate = new SemaphoreSlim(Math.Max(1, job.Target.ParallelChunks));
        var sends = job.Chunks.Select(async chunk =>
        {
            await gate.WaitAsync(cancellationToken);
            try
            {
                await SendChunkAsync(job, chunk, partsDirectory, remoteParts, cancellationToken);
            }
            finally
            {
                gate.Release();
            }
        });
        await Task.WhenAll(sends);

        var firstFailure = job.Chunks.FirstOrDefault(chunk => chunk.State == ChunkState.Failed);
        if (firstFailure is not null)
        {
            throw new InvalidOperationException(
                $"One or more chunks could not be sent. {firstFailure.Error}");
        }

        await RunRequiredAsync(
            Rclone,
            [
                "copyto",
                Path.Combine(partsDirectory, ManifestFileName),
                $"{remoteParts}/{ManifestFileName}",
                "--retries", "2",
                "--low-level-retries", "10",
                "--timeout", "10m",
                "--contimeout", "30s",
                "--sftp-disable-hashcheck",
            ],
            cancellationToken);
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

            var result = await RunAsync(
                Rclone,
                [
                    "copyto", localPath, $"{remoteParts}/{chunk.FileName}",
                    "--retries", "2",
                    "--low-level-retries", "10",
                    "--timeout", "10m",
                    "--contimeout", "30s",
                    "--sftp-disable-hashcheck",
                ],
                cancellationToken);
            if (result.Succeeded)
            {
                chunk.State = ChunkState.Sent;
                chunk.Error = null;
                Report(job);
                return;
            }

            chunk.Error = LastUsefulLine(result.CombinedOutput) ??
                $"rclone exited with {result.ExitCode}.";
            if (attempt < maxAttempts)
            {
                await Task.Delay(TimeSpan.FromSeconds(2 * attempt), cancellationToken);
            }
        }

        chunk.State = ChunkState.Failed;
        Report(job);
    }

    private async Task AssembleAsync(UploadJob job, CancellationToken cancellationToken)
    {
        job.Phase = UploadPhase.Assembling;
        Report(job);

        var remoteParts = CombineRemoteWindows(
            job.Target.RemotePartsRoot,
            job.Workspace,
            job.TransferDirectoryName);
        var result = await SshAsync(
            job.Target,
            EncodePowerShell(BuildRemoteScript(remoteParts, job.ManifestSha256)),
            cancellationToken);
        if (!result.Succeeded)
        {
            throw new InvalidOperationException(
                $"Remote assembly failed: {result.CombinedOutput.Trim()}");
        }

        job.RemoteSha256 = result.StandardOutput
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .FirstOrDefault(line => line.StartsWith(
                "MEDIA_PIPELINE_VERIFIED:", StringComparison.OrdinalIgnoreCase))?
            .Split(':', 2)[1];
    }

    private Task VerifyAsync(UploadJob job, CancellationToken cancellationToken)
    {
        job.Phase = UploadPhase.Verifying;
        Report(job);

        if (!string.Equals(job.RemoteSha256, job.SourceSha256, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "The remote worker did not verify the uploaded source hash.");
        }

        job.RemoteVerified = true;
        return Task.CompletedTask;
    }

    private static async Task<ProcessResult> SshAsync(
        UploadTarget target,
        string command,
        CancellationToken cancellationToken)
    {
        var arguments = new List<string>
        {
            "-o", "BatchMode=yes",
            "-o", "ConnectTimeout=8",
            "-o", "ServerAliveInterval=30",
            "-o", "ServerAliveCountMax=3",
            "-o", "TCPKeepAlive=yes",
        };
        if (!string.IsNullOrWhiteSpace(target.SshKeyFile))
        {
            arguments.AddRange(["-i", target.SshKeyFile]);
        }

        arguments.AddRange([
            "-p", target.SshPort.ToString(System.Globalization.CultureInfo.InvariantCulture),
            target.SshHost,
            command,
        ]);
        return await RunAsync(Ssh, arguments, cancellationToken);
    }

    private static string EncodePowerShell(string script) =>
        $"powershell -NoProfile -NonInteractive -EncodedCommand " +
        Convert.ToBase64String(Encoding.Unicode.GetBytes(script));

    private static string EscapePowerShellLiteral(string value) => value.Replace("'", "''");

    private static string CombineRemoteWindows(params string[] parts) =>
        string.Join('\\', parts
            .Select((part, index) => index == 0
                ? part.Replace('/', '\\').TrimEnd('\\')
                : part.Replace('/', '\\').Trim('\\')));

    private static async Task<ProcessResult> RunAsync(
        string executable,
        IEnumerable<string> arguments,
        CancellationToken cancellationToken)
    {
        try
        {
            return await ProcessRunner.RunAsync(executable, arguments, cancellationToken);
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            throw new InvalidOperationException(
                $"Could not run {executable}. Install it and make it available on PATH. " +
                exception.Message,
                exception);
        }
    }

    private static async Task RunRequiredAsync(
        string executable,
        IEnumerable<string> arguments,
        CancellationToken cancellationToken)
    {
        var result = await RunAsync(executable, arguments, cancellationToken);
        if (!result.Succeeded)
        {
            throw new InvalidOperationException(
                $"{executable} exited with {result.ExitCode}: {result.CombinedOutput.Trim()}");
        }
    }

    private static string? LastUsefulLine(string output) => output
        .Split(new[] { '\r', '\n' },
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
        .LastOrDefault();

    private FileStream AcquireUploadLease(UploadJob job)
    {
        var partsDirectory = LocalPartsDirectory(job);
        Directory.CreateDirectory(partsDirectory);
        try
        {
            return new FileStream(
                LocalLeasePath(job),
                FileMode.OpenOrCreate,
                FileAccess.ReadWrite,
                FileShare.None);
        }
        catch (IOException exception)
        {
            throw new InvalidOperationException(
                $"Another upload already owns '{job.Workspace}/{job.FileName}'.",
                exception);
        }
    }

    private static async Task EnsureSourceUnchangedAsync(
        UploadJob job,
        CancellationToken cancellationToken)
    {
        if (!File.Exists(job.SourcePath) || new FileInfo(job.SourcePath).Length != job.TotalBytes ||
            !string.Equals(
                await FileChunker.HashAsync(job.SourcePath, cancellationToken),
                job.SourceSha256,
                StringComparison.OrdinalIgnoreCase))
        {
            throw new IOException("The upload source changed before remote verification completed.");
        }
    }

    private void Cleanup(UploadJob job)
    {
        try
        {
            var partsDirectory = LocalPartsDirectory(job);
            if (Directory.Exists(partsDirectory))
            {
                Directory.Delete(partsDirectory, recursive: true);
            }
        }
        catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException)
        {
        }
    }

    internal static async Task ClaimSourceForDeletionAsync(
        UploadJob job,
        CancellationToken cancellationToken = default)
    {
        if (!job.RemoteVerified)
        {
            throw new InvalidOperationException("Cannot claim an upload source before remote verification.");
        }

        var directory = Path.GetDirectoryName(job.SourcePath)
            ?? throw new InvalidOperationException($"No parent directory for '{job.SourcePath}'.");
        var claimPath = Path.Combine(
            directory,
            $".{job.FileName}.media-pipeline-delete-{Guid.NewGuid():N}.claim");
        File.Move(job.SourcePath, claimPath);
        job.ClaimedSourcePath = claimPath;

        try
        {
            job.ClaimedSourceLease = new FileStream(
                claimPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.None,
                bufferSize: 1 << 20,
                useAsync: true);
            if (job.ClaimedSourceLease.Length != job.TotalBytes ||
                !string.Equals(
                    Convert.ToHexString(await SHA256.HashDataAsync(
                        job.ClaimedSourceLease,
                        cancellationToken)),
                    job.SourceSha256,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new IOException(
                    "The upload source was replaced before it could be claimed for deletion.");
            }
        }
        catch
        {
            RestoreClaimedSource(job);
            throw;
        }
    }

    internal static void DeleteClaimedSource(UploadJob job)
    {
        if (!job.RemoteVerified || job.ClaimedSourcePath is null)
        {
            return;
        }

        // Windows sharing rules let the held lease prove that no writer still owns the inode.
        // POSIX permits an unrelated process to keep writing through a descriptor after rename,
        // so the macOS app moves the verified claim to Trash through Foundation instead.
        if (!OperatingSystem.IsWindows())
        {
            job.ClaimedSourceLease?.Dispose();
            job.ClaimedSourceLease = null;
            job.RetainedSourcePath = job.ClaimedSourcePath;
            job.ClaimedSourcePath = null;
            return;
        }

        try
        {
            job.ClaimedSourceLease?.Dispose();
            job.ClaimedSourceLease = null;
            File.Delete(job.ClaimedSourcePath);
            job.ClaimedSourcePath = null;
            job.SourceDeleted = true;
        }
        catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException)
        {
            RestoreClaimedSource(job);
        }
    }

    private static void RestoreClaimedSource(UploadJob job)
    {
        job.ClaimedSourceLease?.Dispose();
        job.ClaimedSourceLease = null;
        var claimPath = job.ClaimedSourcePath;
        if (claimPath is null || !File.Exists(claimPath))
        {
            job.ClaimedSourcePath = null;
            return;
        }
        job.RetainedSourcePath = claimPath;

        var directory = Path.GetDirectoryName(job.SourcePath)!;
        var extension = Path.GetExtension(job.FileName);
        var stem = Path.GetFileNameWithoutExtension(job.FileName);
        for (var attempt = 0; attempt < 100; attempt++)
        {
            var destination = attempt == 0
                ? job.SourcePath
                : Path.Combine(directory, $"{stem}.upload-retry-{attempt}{extension}");
            try
            {
                File.Move(claimPath, destination);
                job.ClaimedSourcePath = null;
                job.RetainedSourcePath = destination;
                return;
            }
            catch (IOException) when (File.Exists(destination))
            {
            }
            catch (Exception exception) when (
                exception is IOException or UnauthorizedAccessException)
            {
                return;
            }
        }
    }

    private static string DestinationLock(UploadJob job)
    {
        var destination = CombineRemoteWindows(job.RemoteWorkspaceDirectory, job.FileName);
        return Convert.ToHexString(SHA256.HashData(
            Encoding.UTF8.GetBytes(destination.ToLowerInvariant())));
    }

    public static string BuildRemoteScript(string remotePartsDirectory, string manifestSha256)
    {
        return """
        $ErrorActionPreference = 'Stop'

        function Get-Sha256([string]$Path) {
            $stream = [IO.File]::OpenRead($Path)
            $hasher = [Security.Cryptography.SHA256]::Create()
            try { return [BitConverter]::ToString($hasher.ComputeHash($stream)).Replace('-', '') }
            finally {
                $hasher.Dispose()
                $stream.Dispose()
            }
        }

        $partsRoot = '__REMOTE_PARTS__'
        $manifestPath = Join-Path $partsRoot 'manifest.json'
        if (-not (Test-Path -LiteralPath $manifestPath)) {
            throw 'The authenticated upload manifest is missing.'
        }
        $manifestBytes = [IO.File]::ReadAllBytes($manifestPath)
        $sha256 = [Security.Cryptography.SHA256]::Create()
        try { $manifestHash = [BitConverter]::ToString($sha256.ComputeHash($manifestBytes)).Replace('-', '') }
        finally { $sha256.Dispose() }
        if ($manifestHash -ne '__MANIFEST_SHA256__') {
            throw 'The upload manifest failed its checksum.'
        }
        $manifestJson = [Text.Encoding]::UTF8.GetString($manifestBytes)
        $manifest = $manifestJson | ConvertFrom-Json
        $finalPath = Join-Path $manifest.remoteDirectory $manifest.fileName
        $tempPath = "$finalPath.$($manifest.sourceSha256).chunked.tmp"
        $lockDirectory = Join-Path $manifest.remoteDirectory '.media-pipeline-locks'

        if (-not (Test-Path -LiteralPath $manifest.remoteDirectory)) {
            New-Item -ItemType Directory -Path $manifest.remoteDirectory -Force | Out-Null
        }
        if (-not (Test-Path -LiteralPath $lockDirectory)) {
            New-Item -ItemType Directory -Path $lockDirectory -Force | Out-Null
        }

        $lockPath = Join-Path $lockDirectory "$($manifest.destinationLock).lock"
        $lockStream = $null
        try {
            $lockStream = [IO.File]::Open(
                $lockPath,
                [IO.FileMode]::OpenOrCreate,
                [IO.FileAccess]::ReadWrite,
                [IO.FileShare]::None)

            if (Test-Path -LiteralPath $finalPath) {
                $existingLength = (Get-Item -LiteralPath $finalPath).Length
                $existingHash = Get-Sha256 $finalPath
                if ($existingLength -ne $manifest.expectedLength -or
                    $existingHash -ne $manifest.sourceSha256) {
                    throw "Destination already exists with different content: $finalPath"
                }
            }
            else {
                $output = [IO.File]::Create($tempPath)
                try {
                    foreach ($part in $manifest.parts) {
                        $partPath = Join-Path $partsRoot $part.name
                        if (-not (Test-Path -LiteralPath $partPath)) {
                            throw "Missing part: $($part.name)"
                        }

                        $actualLength = (Get-Item -LiteralPath $partPath).Length
                        if ($actualLength -ne $part.length) {
                            throw "Part $($part.name) is $actualLength bytes, expected $($part.length)"
                        }

                        $actualHash = Get-Sha256 $partPath
                        if ($actualHash -ne $part.sha256) {
                            throw "Part $($part.name) failed its checksum"
                        }

                        $input = [IO.File]::OpenRead($partPath)
                        try { $input.CopyTo($output) } finally { $input.Dispose() }
                    }
                }
                finally { $output.Dispose() }

                $assembledLength = (Get-Item -LiteralPath $tempPath).Length
                $assembledHash = Get-Sha256 $tempPath
                if ($assembledLength -ne $manifest.expectedLength) {
                    throw "Assembled $assembledLength bytes, expected $($manifest.expectedLength)"
                }
                if ($assembledHash -ne $manifest.sourceSha256) {
                    throw "Assembled file failed its source checksum"
                }

                Move-Item -LiteralPath $tempPath -Destination $finalPath
            }

            Write-Output "MEDIA_PIPELINE_VERIFIED:$($manifest.sourceSha256)"
            Remove-Item -LiteralPath $partsRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        finally {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
            if ($null -ne $lockStream) { $lockStream.Dispose() }
        }
        """
            .Replace("__MANIFEST_SHA256__", manifestSha256, StringComparison.Ordinal)
            .Replace("__REMOTE_PARTS__", EscapePowerShellLiteral(remotePartsDirectory), StringComparison.Ordinal);
    }
}
