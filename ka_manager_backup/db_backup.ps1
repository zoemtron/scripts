# --- НАСТРОЙКИ ---
$FileToUpload = Join-Path $PSScriptRoot 'test.txt'

$RclonePath = Join-Path $PSScriptRoot 'rclone'
$ConfigPath = Join-Path $RclonePath 'config\rclone.conf'
$LogPath = Join-Path $PSScriptRoot 'backup.log'

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$FileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($FileToUpload)
$FileExtension = [System.IO.Path]::GetExtension($FileToUpload)

# Ново име на файла, например: db_backup_20261109_153000.bak
$NewFileName = "${FileNameWithoutExt}_${Timestamp}${FileExtension}"
$RcloneRemote = "drive:KA_MANAGER_BACKUP/$NewFileName"

Start-Transcript -Path $LogPath -Append -ErrorAction Stop | Out-Null

try {
    Write-Host "Качване на '$NewFileName' файл..." -ForegroundColor Cyan

    # Изпълнение на rclone командата за качване
    & "$RclonePath\rclone.exe" copyto "$FileToUpload" "$RcloneRemote" --config "$ConfigPath"
    $RcloneExitCode = $LASTEXITCODE
    if ($RcloneExitCode -eq 0) {
        Write-Host "Успех! Файлът е качен в Google Drive." -ForegroundColor Green
    } else {
        Write-Error "Възникна грешка при качването."
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