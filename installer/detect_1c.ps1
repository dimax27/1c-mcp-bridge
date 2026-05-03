# =============================================================================
#  detect_1c.ps1 — диагностическая утилита.
#  Показывает все найденные платформы 1С и состояние их COM-коннекторов.
#  Запускается отдельно — не нужен мастеру установки, помогает пользователю
#  понять "что у меня есть" до запуска инсталлятора.
# =============================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'

Write-Host "=== Поиск установленных платформ 1С ===" -ForegroundColor Cyan
Write-Host ""

$root = 'HKLM:\SOFTWARE\1C\1Cv8'
if (-not (Test-Path $root)) {
    Write-Host "Раздел реестра $root не найден — платформа 1С не установлена." -ForegroundColor Yellow
    exit 1
}

$found = @()
foreach ($key in Get-ChildItem $root) {
    $version = $key.PSChildName
    if ($version -notmatch '^\d+(\.\d+)+') { continue }

    $props = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
    $path  = $props.InstalledLocation
    if (-not $path) { $path = $props.Path }

    # Major-версия "8.5.1.1150" -> 85
    $parts = $version.Split('.')
    if ($parts.Count -ge 2) {
        $major = [int]("$($parts[0])$($parts[1])")
    } else { continue }

    if ($major -lt 82) { continue }

    $progid = "V$major.COMConnector"
    $type = [Type]::GetTypeFromProgID($progid, $false)
    $registered = ($null -ne $type)

    $dll = if ($path) { Join-Path $path 'bin\comcntr.dll' } else { '' }

    $found += [PSCustomObject]@{
        Version   = $version
        ProgID    = $progid
        Path      = $path
        DllExists = if ($dll -and (Test-Path $dll)) { 'да' } else { 'нет' }
        COMRegistered = if ($registered) { 'да' } else { 'нет' }
    }
}

if ($found.Count -eq 0) {
    Write-Host "Совместимых версий 1С не найдено (нужна 8.2 или новее)." -ForegroundColor Yellow
    exit 1
}

$found | Format-Table -AutoSize

Write-Host ""
Write-Host "=== Python ===" -ForegroundColor Cyan
try {
    $v = & py -3.12 --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "py -3.12: $v"
    } else {
        Write-Host "py -3.12 не отвечает." -ForegroundColor Yellow
    }
} catch {
    Write-Host "py-launcher не установлен — будет скачан вместе с Python." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== Claude Desktop ===" -ForegroundColor Cyan
$claudeConfig = Join-Path $env:APPDATA 'Claude\claude_desktop_config.json'
if (Test-Path $claudeConfig) {
    Write-Host "Конфиг: $claudeConfig"
} else {
    Write-Host "Claude Desktop не запускался ни разу или не установлен." -ForegroundColor Yellow
    Write-Host "Скачать: https://claude.ai/download"
}
