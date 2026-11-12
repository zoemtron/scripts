# --- НАСТРОЙКИ ---
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$FileToUpload
)

if ([string]::IsNullOrWhiteSpace($FileToUpload)) {
    throw 'A file path parameter is required. Pass the path when running the script, for example: .\db_backup.ps1 "C:\test\abc.txt".'
}

if (-not (Test-Path -LiteralPath $FileToUpload -PathType Leaf)) {
    throw "File to upload not found: $FileToUpload"
}

$FileToUpload = (Resolve-Path -LiteralPath $FileToUpload).Path

$RclonePath = Join-Path $PSScriptRoot 'rclone'
$ConfigPath = Join-Path $RclonePath 'config\rclone.conf'
$LogPath = Join-Path $PSScriptRoot 'backup.log'

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$FileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($FileToUpload)
$FileExtension = [System.IO.Path]::GetExtension($FileToUpload)

# Ново име на файла, например: db_backup_20261109_153000.bak
$NewFileName = "${FileNameWithoutExt}_${Timestamp}${FileExtension}"
$RcloneRemote = "drive:KA_MANAGER_BACKUP/$NewFileName"

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
[System.IO.File]::AppendAllText($LogPath, "========== Upload run started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==========$([Environment]::NewLine)", $Utf8LogEncoding)
Write-RunLog -Message "Качване на '$NewFileName' файл..." -Color Cyan

# Изпълнение на rclone командата за качване
$RcloneOutput = & "$RclonePath\rclone.exe" copyto "$FileToUpload" "$RcloneRemote" --config "$ConfigPath" 2>&1 | ForEach-Object {
    Write-RunLog -Message ([string]$_)
}
$RcloneExitCode = $LASTEXITCODE

if ($RcloneExitCode -eq 0) {
    Write-RunLog -Message "Успех! Файлът е качен в Google Drive." -Color Green
} else {
    if ($null -eq $RcloneExitCode) {
        $RcloneExitCode = 1
    }
    Write-RunLog -Message "Възникна грешка при качването." -Color Red
}

Write-RunLog -Message "Upload run finished (exit code: $RcloneExitCode)."
if ($RcloneExitCode -ne 0) {
    exit $RcloneExitCode
}