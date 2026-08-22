<#
.SYNOPSIS
    Migrates the pipeline root from the preset-first layout to the workspace-first layout.

.DESCRIPTION
    Old:  <root>\<preset>\<workspace>\{input,output,original,failed,work}
          <root>\archive\<preset>\<workspace>\output
    New:  <root>\<workspace>\<preset>\{input,output,original,failed,work,archive}
          <root>\sync\<workspace>

    Presets are also renamed so the media type is visible in the folder name.

    Nothing is overwritten: a destination that already holds a file of the same name keeps it,
    and the incoming file is reported as a conflict rather than silently replaced.

    Run with -WhatIf first. It prints exactly what it would do and touches nothing.

.EXAMPLE
    pwsh -File tools\Migrate-FolderLayout.ps1 -WhatIf

.EXAMPLE
    pwsh -File tools\Migrate-FolderLayout.ps1 -DeleteArchives -RemoveObsolete
#>
[CmdletBinding()]
param(
    [string]$PipelineRoot,

    # Print the plan without changing anything.
    [switch]$WhatIf,

    # Delete archived output instead of migrating it.
    [switch]$DeleteArchives,

    # Remove folders the new layout has no use for.
    [switch]$RemoveObsolete
)

$ErrorActionPreference = 'Stop'

if (-not $PipelineRoot) {
    $configPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'config.ini'
    $PipelineRoot = 'D:\MediaPipeline'

    if (Test-Path -LiteralPath $configPath) {
        foreach ($raw in Get-Content -LiteralPath $configPath) {
            $line = $raw.Trim()
            if ($line -match '^\s*PipelineRoot\s*=\s*(.+?)\s*(;.*)?$') {
                $PipelineRoot = $matches[1].Trim('"', "'")
                break
            }
        }
    }
}

if (-not (Test-Path -LiteralPath $PipelineRoot)) {
    throw "Pipeline root not found: $PipelineRoot"
}

$workspaces = @('LC', 'MD', 'YL', 'PL', 'general')

# Old preset folder -> new preset folder. The three mixed-media presets keep neutral names
# because they accept video and images together.
$presetRenames = [ordered]@{
    'default'    = 'bulk'
    'videoclean' = 'video-clean'
    'imageclean' = 'image-clean'
    'images'     = 'image-bulk'
    'sets'       = 'sets'
    'setbatch'   = 'sets-batch'
    'assetstore' = 'asset-store'
    'long'       = 'video-long'
}

$roles = @('input', 'output', 'original', 'failed', 'work')

$moved = 0
$conflicts = New-Object System.Collections.Generic.List[string]
$deleted = 0

function Show { param([string]$Text) Write-Host $Text }

function Move-Tree {
    param([string]$Source, [string]$Destination, [string]$Label)

    if (-not (Test-Path -LiteralPath $Source)) { return }

    # -Force throughout: the move below uses it, so the count must see the same files,
    # otherwise a dry run silently under-reports hidden ones such as Thumbs.db.
    $files = @(Get-ChildItem -LiteralPath $Source -Recurse -File -Force -ErrorAction SilentlyContinue)
    $dirs = @(Get-ChildItem -LiteralPath $Source -Directory -Force -ErrorAction SilentlyContinue)

    if ($files.Count -eq 0 -and $dirs.Count -eq 0) { return }

    Show ("  {0,-46} {1} file(s)" -f $Label, $files.Count)

    if ($WhatIf) { return }

    if (-not (Test-Path -LiteralPath $Destination)) {
        New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    }

    # Move top-level entries so folder-grouped output (sets, batches) stays intact.
    foreach ($entry in (Get-ChildItem -LiteralPath $Source -Force)) {
        $target = Join-Path $Destination $entry.Name

        if (Test-Path -LiteralPath $target) {
            $conflicts.Add("$($entry.FullName) -> $target") | Out-Null
            continue
        }

        Move-Item -LiteralPath $entry.FullName -Destination $target -Force
        $script:moved++
    }
}

Show ""
Show "Pipeline root: $PipelineRoot"
Show $(if ($WhatIf) { 'MODE: dry run, nothing will be changed' } else { 'MODE: applying changes' })
Show ""

# --- 1. archives -----------------------------------------------------------

$archiveRoot = Join-Path $PipelineRoot 'archive'

if (Test-Path -LiteralPath $archiveRoot) {
    $archiveFiles = @(Get-ChildItem -LiteralPath $archiveRoot -Recurse -File -Force -ErrorAction SilentlyContinue)

    if ($DeleteArchives) {
        Show "Deleting archived output ($($archiveFiles.Count) file(s))"

        if (-not $WhatIf) {
            Remove-Item -LiteralPath $archiveRoot -Recurse -Force
            $deleted += $archiveFiles.Count
        }
    }
    else {
        Show "Migrating archived output ($($archiveFiles.Count) file(s))"

        foreach ($oldPreset in $presetRenames.Keys) {
            foreach ($workspace in $workspaces) {
                $source = Join-Path (Join-Path (Join-Path $archiveRoot $oldPreset) $workspace) 'output'
                $destination = Join-Path (Join-Path (Join-Path $PipelineRoot $workspace) $presetRenames[$oldPreset]) 'archive'
                Move-Tree -Source $source -Destination $destination -Label "archive\$oldPreset\$workspace"
            }
        }
    }

    Show ""
}

# --- 2. lanes ---------------------------------------------------------------

Show "Migrating lane folders"

foreach ($oldPreset in $presetRenames.Keys) {
    $newPreset = $presetRenames[$oldPreset]

    foreach ($workspace in $workspaces) {
        foreach ($role in $roles) {
            $source = Join-Path (Join-Path (Join-Path $PipelineRoot $oldPreset) $workspace) $role
            $destination = Join-Path (Join-Path (Join-Path $PipelineRoot $workspace) $newPreset) $role
            Move-Tree -Source $source -Destination $destination -Label "$oldPreset\$workspace\$role"
        }
    }
}

Show ""

# --- 3. sync ----------------------------------------------------------------

# Uploads used to sit in one flat folder. Names already carry a workspace prefix, so they sort
# into workspace folders by that prefix; anything unrecognised goes to the default workspace
# rather than being dropped.
$syncRoot = Join-Path $PipelineRoot 'sync'

if (Test-Path -LiteralPath $syncRoot) {
    Show "Sorting staged uploads into workspace folders"

    foreach ($file in (Get-ChildItem -LiteralPath $syncRoot -File -Force -ErrorAction SilentlyContinue)) {
        $workspace = 'LC'
        foreach ($candidate in $workspaces) {
            if ($file.Name -like "$candidate-*") { $workspace = $candidate; break }
        }

        $destination = Join-Path $syncRoot $workspace
        Show ("  {0,-46} -> sync\{1}" -f $file.Name, $workspace)

        if (-not $WhatIf) {
            if (-not (Test-Path -LiteralPath $destination)) {
                New-Item -ItemType Directory -Path $destination -Force | Out-Null
            }

            $target = Join-Path $destination $file.Name
            if (Test-Path -LiteralPath $target) {
                $conflicts.Add("$($file.FullName) -> $target") | Out-Null
            }
            else {
                Move-Item -LiteralPath $file.FullName -Destination $target -Force
                $script:moved++
            }
        }
    }

    Show ""
}

# --- 4. obsolete folders ----------------------------------------------------

if ($RemoveObsolete) {
    Show "Removing folders the new layout has no use for"

    $obsolete = New-Object System.Collections.Generic.List[string]
    $obsolete.Add((Join-Path $PipelineRoot 'convert')) | Out-Null
    $obsolete.Add((Join-Path $PipelineRoot 'sets_media_temp')) | Out-Null

    foreach ($oldPreset in $presetRenames.Keys) {
        $obsolete.Add((Join-Path (Join-Path $PipelineRoot $oldPreset) 'sync')) | Out-Null
        $obsolete.Add((Join-Path $PipelineRoot $oldPreset)) | Out-Null
    }

    foreach ($directory in $obsolete) {
        if (-not (Test-Path -LiteralPath $directory)) { continue }

        $remaining = @(Get-ChildItem -LiteralPath $directory -Recurse -File -Force -ErrorAction SilentlyContinue)

        if ($remaining.Count -gt 0) {
            Show ("  SKIP {0} still holds {1} file(s)" -f $directory, $remaining.Count)
            continue
        }

        Show "  remove $directory"

        if (-not $WhatIf) {
            Remove-Item -LiteralPath $directory -Recurse -Force
        }
    }

    Show ""
}

# --- summary ----------------------------------------------------------------

Show "----"
if ($WhatIf) {
    Show "Dry run complete. Re-run without -WhatIf to apply."
}
else {
    Show "Moved $moved entr$(if ($moved -eq 1) { 'y' } else { 'ies' }), deleted $deleted archived file(s)."
}

if ($conflicts.Count -gt 0) {
    Show ""
    Show "Conflicts left in place (destination already existed):"
    foreach ($conflict in $conflicts) { Show "  $conflict" }
}
