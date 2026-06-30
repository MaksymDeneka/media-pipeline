[CmdletBinding()]
param(
    [string]$HostName = '100.124.72.13',
    [string]$User = 'root',
    [int]$Port = 2222,
    [string]$KeyFile = (Join-Path $HOME '.ssh\heatup_remote_debug_ed25519'),

    [int]$ClientAliveInterval = 60,
    [int]$ClientAliveCountMax = 3,

    [int]$WatchdogIntervalMinutes = 5,
    [int]$CpuSampleSeconds = 5,
    [double]$CpuPercentOfOneCoreThreshold = 75,
    [int]$MinAgeMinutes = 10,

    [switch]$SkipKeepAliveConfig,
    [switch]$SkipWatchdogTask,
    [switch]$RestartSshd
)

$ErrorActionPreference = 'Stop'

function Invoke-RemotePowerShell {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Script
    )

    $target = "${User}@${HostName}"

    $sshArguments = @(
        '-o', 'BatchMode=yes',
        '-o', 'ConnectTimeout=8',
        '-o', 'ServerAliveInterval=30',
        '-o', 'ServerAliveCountMax=3',
        '-o', 'TCPKeepAlive=yes',
        '-i', $KeyFile,
        '-p', $Port.ToString(),
        $target,
        'powershell -NoProfile -ExecutionPolicy Bypass -Command -'
    )

    $Script | & ssh @sshArguments

    if ($LASTEXITCODE -ne 0) {
        throw "ssh exited with code $LASTEXITCODE"
    }
}

$remoteWatchdog = @'
param(
    [int]$SshPort = 2222,
    [int]$CpuSampleSeconds = 5,
    [double]$CpuPercentOfOneCoreThreshold = 75,
    [int]$MinAgeMinutes = 10,
    [string]$LogPath = 'C:\ProgramData\MediaPipeline\sshd-watchdog.log'
)

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'

function Write-WatchdogLog {
    param([string]$Message)

    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    $directory = Split-Path -Parent $LogPath
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Force -Path $directory | Out-Null
    }

    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
}

function Get-ProcessDescendants {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ProcessId,

        [Parameter(Mandatory = $true)]
        [object[]]$Processes
    )

    $children = @($Processes | Where-Object { [int]$_.ParentProcessId -eq $ProcessId })
    foreach ($child in $children) {
        $child
        Get-ProcessDescendants -ProcessId ([int]$child.ProcessId) -Processes $Processes
    }
}

try {
    $connections = @(Get-NetTCPConnection -LocalPort $SshPort -State Established -ErrorAction SilentlyContinue)
    if ($connections.Count -gt 0) {
        Write-WatchdogLog "Skipped: $($connections.Count) established SSH connection(s) on port $SshPort."
        return
    }

    $service = Get-CimInstance Win32_Service -Filter "Name = 'sshd'" -ErrorAction SilentlyContinue
    $listenerPid = if ($service) { [int]$service.ProcessId } else { 0 }

    $before = @{}
    foreach ($process in @(Get-Process sshd -ErrorAction SilentlyContinue)) {
        $before[[int]$process.Id] = [double]$process.CPU
    }

    Start-Sleep -Seconds $CpuSampleSeconds

    $allProcesses = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue)
    $sshdProcesses = @($allProcesses | Where-Object { $_.Name -ieq 'sshd.exe' })
    $now = Get-Date

    foreach ($candidate in $sshdProcesses) {
        $candidatePid = [int]$candidate.ProcessId
        if ($candidatePid -eq $listenerPid) { continue }
        if ($candidate.CommandLine -notmatch '(^| )-(R|z)( |$)') { continue }

        $liveProcess = Get-Process -Id $candidatePid -ErrorAction SilentlyContinue
        if (-not $liveProcess) { continue }
        if (-not $before.ContainsKey($candidatePid)) { continue }

        $ageMinutes = ($now - $candidate.CreationDate).TotalMinutes
        if ($ageMinutes -lt $MinAgeMinutes) { continue }

        $cpuDelta = [double]$liveProcess.CPU - [double]$before[$candidatePid]
        $percentOfOneCore = ($cpuDelta / [Math]::Max(1, $CpuSampleSeconds)) * 100
        if ($percentOfOneCore -lt $CpuPercentOfOneCoreThreshold) { continue }

        $descendants = @(Get-ProcessDescendants -ProcessId $candidatePid -Processes $allProcesses)
        $nonSshdDescendants = @($descendants | Where-Object { $_.Name -ine 'sshd.exe' })
        if ($nonSshdDescendants.Count -gt 0) {
            Write-WatchdogLog ("Skipped PID {0}: {1:n1}% of one core but has non-sshd descendants: {2}" -f $candidatePid, $percentOfOneCore, (($nonSshdDescendants | ForEach-Object { "$($_.Name):$($_.ProcessId)" }) -join ', '))
            continue
        }

        Write-WatchdogLog ("Stopping stale sshd worker PID {0}: {1:n1}% of one core, age {2:n1} min, command {3}" -f $candidatePid, $percentOfOneCore, $ageMinutes, $candidate.CommandLine)
        Stop-Process -Id $candidatePid -Force
    }
}
catch {
    Write-WatchdogLog "Error: $($_.Exception.Message)"
    throw
}
'@

$remoteWatchdogB64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($remoteWatchdog))
$skipKeepAliveLiteral = if ($SkipKeepAliveConfig) { '$true' } else { '$false' }
$skipWatchdogLiteral = if ($SkipWatchdogTask) { '$true' } else { '$false' }
$restartSshdLiteral = if ($RestartSshd) { '$true' } else { '$false' }

$remoteInstaller = @"
`$ErrorActionPreference = 'Stop'
`$ProgressPreference = 'SilentlyContinue'

`$installRoot = 'C:\ProgramData\MediaPipeline'
`$watchdogPath = Join-Path `$installRoot 'Watch-OpenSshd.ps1'
`$configPath = 'C:\ProgramData\ssh\sshd_config'
`$taskName = 'Media Pipeline SSHD Watchdog'
`$taskPath = '\MediaPipeline\'

New-Item -ItemType Directory -Force -Path `$installRoot | Out-Null

function Set-SshdGlobalSetting {
    param(
        [Parameter(Mandatory = `$true)][string]`$Path,
        [Parameter(Mandatory = `$true)][string]`$Name,
        [Parameter(Mandatory = `$true)][string]`$Value
    )

    if (-not (Test-Path -LiteralPath `$Path)) {
        throw "Missing sshd config: `$Path"
    }

    `$lines = New-Object System.Collections.Generic.List[string]
    foreach (`$line in (Get-Content -LiteralPath `$Path)) {
        `$lines.Add([string]`$line) | Out-Null
    }

    `$matchIndex = -1
    for (`$i = 0; `$i -lt `$lines.Count; `$i++) {
        if (`$lines[`$i] -match '^\s*Match\s+') {
            `$matchIndex = `$i
            break
        }
    }

    `$limit = if (`$matchIndex -ge 0) { `$matchIndex } else { `$lines.Count }
    `$replacement = "`$Name `$Value"
    `$changed = `$false

    for (`$i = 0; `$i -lt `$limit; `$i++) {
        if (`$lines[`$i] -match ("^\s*#?\s*" + [regex]::Escape(`$Name) + "\b")) {
            if (`$lines[`$i] -ne `$replacement) {
                `$lines[`$i] = `$replacement
            }
            `$changed = `$true
            break
        }
    }

    if (-not `$changed) {
        if (`$matchIndex -ge 0) {
            `$lines.Insert(`$matchIndex, `$replacement)
        }
        else {
            `$lines.Add(`$replacement) | Out-Null
        }
    }

    Set-Content -LiteralPath `$Path -Value `$lines -Encoding ascii
}

if (-not $skipKeepAliveLiteral) {
    Copy-Item -LiteralPath `$configPath -Destination ("`$configPath.media-pipeline-backup-" + (Get-Date -Format 'yyyyMMdd-HHmmss')) -Force
    Set-SshdGlobalSetting -Path `$configPath -Name 'TCPKeepAlive' -Value 'yes'
    Set-SshdGlobalSetting -Path `$configPath -Name 'ClientAliveInterval' -Value '$ClientAliveInterval'
    Set-SshdGlobalSetting -Path `$configPath -Name 'ClientAliveCountMax' -Value '$ClientAliveCountMax'
    Write-Output "Updated sshd keepalive settings in `$configPath."
}
else {
    Write-Output 'Skipped sshd keepalive config update.'
}

if (-not $skipWatchdogLiteral) {
    `$watchdogText = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$remoteWatchdogB64'))
    [IO.File]::WriteAllText(`$watchdogPath, `$watchdogText, [Text.UTF8Encoding]::new(`$false))

    `$taskArgument = '-NoProfile -ExecutionPolicy Bypass -File "' + `$watchdogPath + '" -SshPort $Port -CpuSampleSeconds $CpuSampleSeconds -CpuPercentOfOneCoreThreshold $CpuPercentOfOneCoreThreshold -MinAgeMinutes $MinAgeMinutes'
    `$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument `$taskArgument
    `$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes $WatchdogIntervalMinutes) -RepetitionDuration (New-TimeSpan -Days 3650)
    `$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 2) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    `$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -RunLevel Highest

    Register-ScheduledTask -TaskName `$taskName -TaskPath `$taskPath -Action `$action -Trigger `$trigger -Settings `$settings -Principal `$principal -Force | Out-Null
    Write-Output "Installed watchdog script at `$watchdogPath."
    Write-Output "Registered scheduled task `$taskPath`$taskName every $WatchdogIntervalMinutes minute(s)."
}
else {
    Write-Output 'Skipped watchdog scheduled task install.'
}

if ($restartSshdLiteral) {
    `$restartTaskName = 'OneShotRestartSshd'
    `$restartAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -ExecutionPolicy Bypass -Command "Restart-Service sshd -Force"'
    `$restartTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddDays(1)
    `$restartSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 2)
    `$restartPrincipal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -RunLevel Highest
    Register-ScheduledTask -TaskName `$restartTaskName -TaskPath `$taskPath -Action `$restartAction -Trigger `$restartTrigger -Settings `$restartSettings -Principal `$restartPrincipal -Force | Out-Null
    Start-ScheduledTask -TaskName `$restartTaskName -TaskPath `$taskPath
    Start-Sleep -Seconds 1
    Disable-ScheduledTask -TaskName `$restartTaskName -TaskPath `$taskPath | Out-Null
    Write-Output 'Started one-shot SYSTEM task to restart sshd, then disabled its placeholder trigger.'
}
elseif (-not $skipKeepAliveLiteral) {
    Write-Output 'Keepalive settings will apply after the sshd service is restarted.'
}

Get-Service sshd | Select-Object Name,Status,StartType | Format-Table -AutoSize | Out-String -Width 200 | Write-Output
"@

Write-Host "Installing SSHD protection on ${User}@${HostName}:$Port"
Write-Host "Keepalive config:     $(-not $SkipKeepAliveConfig)"
Write-Host "Watchdog task:        $(-not $SkipWatchdogTask)"
Write-Host "Restart sshd now:     $($RestartSshd.IsPresent)"
Write-Host "Watchdog threshold:   $CpuPercentOfOneCoreThreshold% of one core for $CpuSampleSeconds seconds"
Write-Host "Minimum worker age:   $MinAgeMinutes minutes"

Invoke-RemotePowerShell -Script $remoteInstaller
