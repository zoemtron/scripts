# --- НАСТРОЙКИ ---
$PSScriptRootPath = "C:\ZOEMTRON\scripts\ka_manager_backup"
$FileToUpload = "$PSScriptRootPath\test.txt"
$RcloneRemote = "drive:KA_MANAGER_BACKUP"
$RclonePath = "$PSScriptRootPath\rclone-v1.76.0-windows-amd64"

Write-Host "Качване на тестовия файл..." -ForegroundColor Cyan

# Изпълнение на rclone командата за качване
& "$RclonePath\rclone.exe" copy "$FileToUpload" "$RcloneRemote" --config "$RclonePath\config\rclone.conf" -v

if ($LASTEXITCODE -eq 0) {
    Write-Host "Успех! Файлът е качен в Google Drive." -ForegroundColor Green
} else {
    Write-Error "Възникна грешка при качването."
}