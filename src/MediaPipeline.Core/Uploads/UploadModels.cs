using MediaPipeline.Core.Configuration;
using System.Security.Cryptography;
using System.Text;

namespace MediaPipeline.Core.Uploads;

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
}

public sealed class ChunkProgress
{
    public required int Index { get; init; }
    public required long Length { get; init; }
    public required string FileName { get; init; }
    public string Sha256 { get; set; } = "";
    public ChunkState State { get; set; }
    public int Attempts { get; set; }
    public string? Error { get; set; }
}

public sealed record UploadTarget
{
    public required string RemoteName { get; init; }
    public required string RemoteSftpPartsRoot { get; init; }
    public required string RemotePartsRoot { get; init; }
    public required string RemoteDirectory { get; init; }
    public required string SshHost { get; init; }
    public required int SshPort { get; init; }
    public required string SshKeyFile { get; init; }
    public required int ChunkSizeMB { get; init; }
    public required int ParallelChunks { get; init; }
    public required bool DeleteAfterUpload { get; init; }

    public static UploadTarget FromConfiguration(UploadOptions options) => new()
    {
        RemoteName = options.RemoteName,
        RemoteSftpPartsRoot = options.RemoteSftpPartsRoot,
        RemotePartsRoot = options.RemotePartsRoot,
        RemoteDirectory = options.RemoteDirectory,
        SshHost = options.RemoteSshHost,
        SshPort = options.RemoteSshPort,
        SshKeyFile = options.RemoteSshKeyFile,
        ChunkSizeMB = options.ChunkSizeMB,
        ParallelChunks = options.ParallelChunks,
        DeleteAfterUpload = options.DeleteAfterUpload,
    };
}

public sealed class UploadJob
{
    public required string SourcePath { get; init; }
    public required UploadTarget Target { get; init; }
    public string? WorkspaceOverride { get; init; }
    public string SourceSha256 { get; set; } = "";
    public string ManifestSha256 { get; set; } = "";
    public string? RemoteSha256 { get; set; }
    internal string? ClaimedSourcePath { get; set; }
    internal FileStream? ClaimedSourceLease { get; set; }
    public string? RetainedSourcePath { get; internal set; }

    public string FileName => Path.GetFileName(SourcePath);

    public string Workspace
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(WorkspaceOverride))
            {
                return WorkspaceOverride;
            }

            var syncFolder = Path.GetDirectoryName(SourcePath);
            var workspace = Path.GetFileName(Path.GetDirectoryName(syncFolder) ?? "");
            return string.IsNullOrWhiteSpace(workspace) ? "general" : workspace;
        }
    }

    public string RemoteWorkspaceDirectory =>
        $"{Target.RemoteDirectory.TrimEnd('\\', '/')}\\{Workspace}";

    public string TransferDirectoryName
    {
        get
        {
            var nameHash = Convert.ToHexString(
                SHA256.HashData(Encoding.UTF8.GetBytes(FileName.ToLowerInvariant())))[..16]
                .ToLowerInvariant();
            return $"{nameHash}.{SourceSha256.ToLowerInvariant()}.parts";
        }
    }

    public long TotalBytes { get; set; }
    public List<ChunkProgress> Chunks { get; } = [];
    public UploadPhase Phase { get; set; }
    public string? Error { get; set; }
    public DateTimeOffset? StartedUtc { get; set; }
    public bool RemoteVerified { get; set; }
    public bool SourceDeleted { get; set; }

    public int ChunksSent => Chunks.Count(chunk => chunk.State == ChunkState.Sent);
    public long BytesSent => Chunks
        .Where(chunk => chunk.State == ChunkState.Sent)
        .Sum(chunk => chunk.Length);
    public double Fraction => TotalBytes > 0 ? (double)BytesSent / TotalBytes : 0;
}
