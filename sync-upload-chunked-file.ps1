[CmdletBinding()]
param(
    [string]$FilePath,
    [string]$LocalPipelineRoot = 'D:\MediaPipeline',
    [string]$RemoteName = 'heatup-remote-sftp',
    [string]$RemoteSftpPartsRoot = '/D:/MediaPipeline/.sync-parts',
    [string]$RemotePartsRoot = 'D:\MediaPipeline\.sync-parts',
    [string]$RemoteDirectory = 'D:\MediaPipeline\sync',
    [int]$ChunkSizeMB = 256,
    [int]$Transfers = 12,
    [switch]$SkipSplit,
    [switch]$NoAssemble
)

$ErrorActionPreference = 'Stop'

function Invoke-Checked {
    param(
        [string]$Command,
        [string[]]$Arguments
    )

    & $Command @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Command exited with code $LASTEXITCODE"
    }
}

function Get-DefaultUploadFile {
    param([string]$SyncRoot)

    $candidate = Get-ChildItem -Path $SyncRoot -File |
        Where-Object { $_.Name -notlike '*.rclone-partial' -and $_.Name -notlike '*.chunked.tmp' } |
        Sort-Object Length -Descending |
        Select-Object -First 1

    if (-not $candidate) {
        throw "No files found in $SyncRoot"
    }

    return $candidate.FullName
}

function Split-FileIntoParts {
    param(
        [System.IO.FileInfo]$Source,
        [string]$PartsDirectory,
        [int64]$ChunkSizeBytes,
        [int]$ChunkCount
    )

    New-Item -ItemType Directory -Force -Path $PartsDirectory | Out-Null

    $buffer = New-Object byte[] (4MB)
    $inputStream = [System.IO.File]::Open($Source.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)

    try {
        for ($chunkIndex = 0; $chunkIndex -lt $ChunkCount; $chunkIndex++) {
            $partName = '{0}.part{1:D5}' -f $Source.Name, ($chunkIndex + 1)
            $partPath = Join-Path $PartsDirectory $partName
            $offset = [int64]$chunkIndex * $ChunkSizeBytes
            $expectedBytes = [Math]::Min($ChunkSizeBytes, $Source.Length - $offset)

            if ((Test-Path $partPath) -and ((Get-Item $partPath).Length -eq $expectedBytes)) {
                Write-Host ("Part {0}/{1} already exists: {2}" -f ($chunkIndex + 1), $ChunkCount, $partName)
                continue
            }

            Write-Host ("Writing part {0}/{1}: {2}" -f ($chunkIndex + 1), $ChunkCount, $partName)
            $inputStream.Seek($offset, [System.IO.SeekOrigin]::Begin) | Out-Null
            $outputStream = [System.IO.File]::Open($partPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)

            try {
                $remaining = $expectedBytes
                while ($remaining -gt 0) {
                    $readSize = [Math]::Min($buffer.Length, $remaining)
                    $bytesRead = $inputStream.Read($buffer, 0, $readSize)
                    if ($bytesRead -le 0) {
                        throw "Unexpected end of source file while writing $partName"
                    }

                    $outputStream.Write($buffer, 0, $bytesRead)
                    $remaining -= $bytesRead
                }
            }
            finally {
                $outputStream.Dispose()
            }
        }
    }
    finally {
        $inputStream.Dispose()
    }
}

$syncRoot = Join-Path $LocalPipelineRoot 'sync'
if ([string]::IsNullOrWhiteSpace($FilePath)) {
    $FilePath = Get-DefaultUploadFile -SyncRoot $syncRoot
}

$source = Get-Item -LiteralPath $FilePath
if (-not $source.PSIsContainer -and $source.Length -le 0) {
    throw "Source file is empty: $($source.FullName)"
}

$chunkSizeBytes = [int64]$ChunkSizeMB * 1MB
$chunkCount = [int][Math]::Ceiling($source.Length / $chunkSizeBytes)
$localPartsRoot = Join-Path $LocalPipelineRoot '.sync-parts'
$partsDirectory = Join-Path $localPartsRoot ($source.Name + '.parts')
$remotePartsPath = ($RemoteSftpPartsRoot.TrimEnd('/') + '/' + ($source.Name + '.parts'))
$remoteRclonePath = "${RemoteName}:$remotePartsPath"

Write-Host "Chunked source:       $($source.FullName)"
Write-Host "File size:            $([Math]::Round($source.Length / 1GB, 3)) GiB"
Write-Host "Chunk size:           $ChunkSizeMB MB"
Write-Host "Chunk count:          $chunkCount"
Write-Host "Local parts folder:   $partsDirectory"
Write-Host "Remote parts folder:  $remoteRclonePath"
Write-Host "Parallel transfers:   $Transfers"

if (-not $SkipSplit) {
    Split-FileIntoParts -Source $source -PartsDirectory $partsDirectory -ChunkSizeBytes $chunkSizeBytes -ChunkCount $chunkCount
}

$manifest = [ordered]@{
    fileName = $source.Name
    expectedLength = $source.Length
    chunkSizeBytes = $chunkSizeBytes
    chunkCount = $chunkCount
    remoteDirectory = $RemoteDirectory
    remotePartsDirectory = (Join-Path $RemotePartsRoot ($source.Name + '.parts'))
    sourceLastWriteTimeUtc = $source.LastWriteTimeUtc.ToString('o')
}

$manifestPath = Join-Path $partsDirectory 'manifest.json'
$manifest | ConvertTo-Json -Depth 3 | Set-Content -Path $manifestPath -Encoding UTF8

Invoke-Checked -Command 'rclone' -Arguments @(
    'mkdir',
    $remoteRclonePath,
    '--timeout', '30s'
)

Invoke-Checked -Command 'rclone' -Arguments @(
    'copy',
    $partsDirectory,
    $remoteRclonePath,
    '--progress',
    '--stats', '10s',
    '--transfers', $Transfers.ToString(),
    '--checkers', '16',
    '--retries', '5',
    '--low-level-retries', '20',
    '--timeout', '10m',
    '--contimeout', '30s',
    '--size-only',
    '--sftp-disable-hashcheck'
)

if ($NoAssemble) {
    Write-Host 'Skipping remote reassembly because -NoAssemble was passed.'
    exit 0
}

$payloadJson = $manifest | ConvertTo-Json -Compress
$payloadB64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($payloadJson))

$remoteScript = @"
`$ErrorActionPreference = 'Stop'
`$payloadJson = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$payloadB64'))
`$payload = `$payloadJson | ConvertFrom-Json
`$partsDirectory = [string]`$payload.remotePartsDirectory
`$remoteDirectory = [string]`$payload.remoteDirectory
`$fileName = [string]`$payload.fileName
`$expectedLength = [int64]`$payload.expectedLength
`$chunkCount = [int]`$payload.chunkCount
`$sourceLastWriteTimeUtc = [DateTime]::Parse([string]`$payload.sourceLastWriteTimeUtc)

New-Item -ItemType Directory -Force -Path `$remoteDirectory | Out-Null
`$finalPath = Join-Path `$remoteDirectory `$fileName
`$tmpPath = Join-Path `$remoteDirectory (`$fileName + '.chunked.tmp')
`$parts = Get-ChildItem -Path `$partsDirectory -Filter (`$fileName + '.part*') -File | Sort-Object Name

if (`$parts.Count -ne `$chunkCount) {
    throw "Expected `$chunkCount parts but found `$(`$parts.Count) in `$partsDirectory"
}

`$buffer = New-Object byte[] (4MB)
`$outStream = [System.IO.File]::Open(`$tmpPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
try {
    foreach (`$part in `$parts) {
        Write-Host ("Appending " + `$part.Name)
        `$inStream = [System.IO.File]::Open(`$part.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        try {
            while ((`$bytesRead = `$inStream.Read(`$buffer, 0, `$buffer.Length)) -gt 0) {
                `$outStream.Write(`$buffer, 0, `$bytesRead)
            }
        }
        finally {
            `$inStream.Dispose()
        }
    }
}
finally {
    `$outStream.Dispose()
}

`$actualLength = (Get-Item -LiteralPath `$tmpPath).Length
if (`$actualLength -ne `$expectedLength) {
    throw "Reassembled file length `$actualLength did not match expected `$expectedLength"
}

Move-Item -LiteralPath `$tmpPath -Destination `$finalPath -Force
(Get-Item -LiteralPath `$finalPath).LastWriteTimeUtc = `$sourceLastWriteTimeUtc
Write-Host "Reassembled `$finalPath (`$actualLength bytes)"
"@

$encodedRemoteScript = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($remoteScript))
Invoke-Checked -Command 'ssh' -Arguments @(
    '-o', 'BatchMode=yes',
    '-o', 'ConnectTimeout=8',
    '-i', (Join-Path $HOME '.ssh\heatup_remote_debug_ed25519'),
    '-p', '2222',
    'root@100.124.72.13',
    "powershell -NoProfile -EncodedCommand $encodedRemoteScript"
)
