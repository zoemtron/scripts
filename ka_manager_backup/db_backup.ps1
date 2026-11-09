# --- НАСТРОЙКИ ---
$PSScriptRootPath = "C:\ZOEMTRON\scripts\ka_manager_backup"
$FileToUpload = "$PSScriptRootPath\test.txt"

$RclonePath = "$PSScriptRootPath\rclone"
$ConfigPath = "$RclonePath\config\rclone.conf"

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$FileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($FileToUpload)
$FileExtension = [System.IO.Path]::GetExtension($FileToUpload)

# Ново име на файла, например: db_backup_20261109_153000.bak
$NewFileName = "${FileNameWithoutExt}_${Timestamp}${FileExtension}"
$RcloneRemote = "drive:KA_MANAGER_BACKUP/$NewFileName"

Write-Host "Качване на '$NewFileName' файл..." -ForegroundColor Cyan

# Изпълнение на rclone командата за качване
& "$RclonePath\rclone.exe" copyto "$FileToUpload" "$RcloneRemote" --config "$ConfigPath"
if ($LASTEXITCODE -eq 0) {
    Write-Host "Успех! Файлът е качен в Google Drive." -ForegroundColor Green
} else {
    Write-Error "Възникна грешка при качването."
}