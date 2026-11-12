# Път до rclone и конфигурацията
[CmdletBinding()]
param(
    [ValidateRange(1, 36500)]
    [int]$RetentionDays = 10
)

$RclonePath = Join-Path $PSScriptRoot 'rclone'
$ConfigPath = Join-Path $RclonePath 'config\rclone.conf'
$LogPath = Join-Path $PSScriptRoot 'backup.log'
$RemoteTarget = "drive:KA_MANAGER_BACKUP"

if (Test-Path $LogPath) {
    $LogContents = [System.IO.File]::ReadAllText($LogPath)
    if ($LogContents.Length -gt 0) {
        $PreviousRuns = @([regex]::Split($LogContents, '(?m)(?=^(?:Windows PowerShell transcript start|========== (?:Upload|Cleanup) run started:))') | Where-Object { $_.Trim().Length -gt 0 })
        $RetainedRuns = ($PreviousRuns | Select-Object -Last 9) -join ''
        [System.IO.File]::WriteAllText($LogPath, $RetainedRuns, [System.Text.UTF8Encoding]::new($false))
    }
}

function Write-RunLog {
    param(
        [string]$Message,
        [System.ConsoleColor]$Color = [System.ConsoleColor]::Gray
    )

    Write-Host $Message -ForegroundColor $Color
    $LogEntry = "[{0}] {1}{2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message, [Environment]::NewLine
    [System.IO.File]::AppendAllText($LogPath, $LogEntry, [System.Text.UTF8Encoding]::new($false))
}

$Utf8LogEncoding = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::AppendAllText($LogPath, "========== Cleanup run started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==========$([Environment]::NewLine)", $Utf8LogEncoding)
Write-RunLog -Message "Стартиране на почистването на архивни файлове по-стари от $RetentionDays дни в $RemoteTarget..." -Color Cyan

$CountOutput = & "$RclonePath\rclone.exe" size $RemoteTarget --json --config $ConfigPath 2>&1

try {
    $CountInfo = ConvertFrom-Json -InputObject ($CountOutput -join [Environment]::NewLine) -ErrorAction Stop
    if ($null -eq $CountInfo.PSObject.Properties['count']) {
        throw 'The count is missing from rclone output.'
    }
    $BackupCount = [long]$CountInfo.count
    if ($BackupCount -lt 0) {
        throw 'Invalid file count.'
    }
} catch {
    foreach ($Line in $CountOutput) {
        Write-RunLog -Message ([string]$Line)
    }
    Write-RunLog -Message "Грешка при преброяването на резервните копия. Няма да бъдат изтрити файлове." -Color Red
    Write-RunLog -Message "Cleanup run finished (exit code: 1)."
    exit 1
}

Write-RunLog -Message "Намерени файлове: $BackupCount."
if ($BackupCount -lt $RetentionDays) {
    Write-RunLog -Message "Файловете са по-малко от $RetentionDays; няма да бъдат изтрити файлове." -Color Green
    Write-RunLog -Message "Cleanup run finished (exit code: 0)."
    exit 0
}

# Изпълнение на rclone delete с ограничение по броя дни за съхранение
$RcloneOutput = & "$RclonePath\rclone.exe" delete $RemoteTarget --min-age "${RetentionDays}d" --config $ConfigPath -v 2>&1 | ForEach-Object {
    Write-RunLog -Message ([string]$_)
}
$RcloneExitCode = $LASTEXITCODE

if ($RcloneExitCode -eq 0) {
    Write-RunLog -Message "Успех! Почистването приключи." -Color Green
} else {
    if ($null -eq $RcloneExitCode) {
        $RcloneExitCode = 1
    }
    Write-RunLog -Message "Възникна грешка при почистването." -Color Red
}

Write-RunLog -Message "Cleanup run finished (exit code: $RcloneExitCode)."
if ($RcloneExitCode -ne 0) {
    exit $RcloneExitCode
}
