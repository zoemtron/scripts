# Път до rclone и конфигурацията
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
Write-RunLog -Message "Стартиране на почистването на архивни файлове по-стари от 10 дни в $RemoteTarget..." -Color Cyan

# Изпълнение на rclone delete с ограничение --min-age 10d
$RcloneOutput = & "$RclonePath\rclone.exe" delete $RemoteTarget --min-age 10d --config $ConfigPath -v 2>&1 | ForEach-Object {
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
