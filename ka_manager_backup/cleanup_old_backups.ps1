# Път до rclone и конфигурацията
$RclonePath = Join-Path $PSScriptRoot 'rclone'
$ConfigPath = Join-Path $RclonePath 'config\rclone.conf'
$LogPath = Join-Path $PSScriptRoot 'backup.log'
$RemoteTarget = "drive:KA_MANAGER_BACKUP"

Start-Transcript -Path $LogPath -Append -ErrorAction Stop | Out-Null

try {
    Write-Host "Стартиране на почистването на архивни файлове по-стари от 10 дни в $RemoteTarget..." -ForegroundColor Cyan

    # Изпълнение на rclone delete с ограничение --min-age 10d
    & "$RclonePath\rclone.exe" delete $RemoteTarget --min-age 10d --config $ConfigPath -v
    $RcloneExitCode = $LASTEXITCODE
    if ($RcloneExitCode -eq 0) {
        Write-Host "Успех! Почистването приключи." -ForegroundColor Green
    } else {
        Write-Error "Възникна грешка при почистването."
    }
} finally {
    Stop-Transcript | Out-Null
}

if ($RcloneExitCode -ne 0) {
    if ($null -eq $RcloneExitCode) {
        exit 1
    }
    exit $RcloneExitCode
}
