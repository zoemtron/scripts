# Път до rclone и конфигурацията
$PSScriptRootPath = "C:\ZOEMTRON\scripts\ka_manager_backup"
$RclonePath = "$PSScriptRootPath\rclone"
$ConfigPath = "$RclonePath\config\rclone.conf"
$RemoteTarget = "drive:KA_MANAGER_BACKUP"

# Съобщение в конзолата
Write-Output "Стартиране на почистване на архивни файлове по-стари от 10 дни в $RemoteTarget..."

# Изпълнение на rclone delete с ограничение --max-age 10d
& "$RclonePath\rclone.exe" delete $RemoteTarget --min-age 10d --config $ConfigPath -v

Write-Host "Почистването приключи."
