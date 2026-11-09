# Път до rclone и конфигурацията
$RclonePath = Join-Path $PSScriptRoot 'rclone'
$ConfigPath = Join-Path $RclonePath 'config\rclone.conf'
$RemoteTarget = "drive:KA_MANAGER_BACKUP"

# Съобщение в конзолата
Write-Host "Стартиране на почистването на архивни файлове по-стари от 10 дни в $RemoteTarget..." -ForegroundColor Cyan

# Изпълнение на rclone delete с ограничение --min-age 10d
& "$RclonePath\rclone.exe" delete $RemoteTarget --min-age 10d --config $ConfigPath -v
$RcloneExitCode = $LASTEXITCODE

if ($RcloneExitCode -eq 0) {
    Write-Host "Успех! Почистването приключи." -ForegroundColor Green
} else {
    Write-Error "Възникна грешка при почистването."
    if ($null -eq $RcloneExitCode) {
        exit 1
    }
    exit $RcloneExitCode
}
