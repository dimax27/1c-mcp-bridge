# =============================================================================
#  restart_http_server.ps1 — перезапуск HTTP-сервера 1C MCP Bridge.
#
#  Останавливает старый сервер (stop_http_server.ps1), запускает VBS-лаунчер
#  заново и проверяет: порт 8000 слушает И настоящий MCP tools/list отвечает.
#  Маркер RESTART_OK выводится ТОЛЬКО при успехе всех проверок.
#
#  Использование:
#    powershell -ExecutionPolicy Bypass -File restart_http_server.ps1 [-Interactive]
# =============================================================================

[CmdletBinding()]
param(
    [switch]$Interactive  # при запуске из ярлыка: пауза после завершения
)

$ErrorActionPreference = 'Continue'

Write-Host "=== Перезапуск HTTP-сервера 1C MCP Bridge ===" -ForegroundColor Cyan

$AppDir = Split-Path $PSScriptRoot -Parent

# 1) Останавливаем старый сервер (точечно: только процесс этой установки)
$StopScript = Join-Path $PSScriptRoot 'stop_http_server.ps1'
if (-not (Test-Path $StopScript)) {
    Write-Host "ОШИБКА: не найден $StopScript" -ForegroundColor Red
    Write-Host "RESTART_FAIL"
    if ($Interactive) { Read-Host "`nНажмите Enter для закрытия..." }
    exit 1
}
. $StopScript
$stopped = Stop-BridgeHttpServer `
    -Port 8000 `
    -ExpectedScriptPath (Join-Path $AppDir 'mcp_server_1c_http.py') `
    -ExpectedPythonPath (Join-Path $AppDir '.venv\Scripts\python.exe')

if (-not $stopped) {
    Write-Host "ОШИБКА: не удалось остановить прежний HTTP-сервер (порт 8000 занят)." -ForegroundColor Red
    Write-Host "RESTART_FAIL"
    if ($Interactive) { Read-Host "`nНажмите Enter для закрытия..." }
    exit 1
}

# 2) Запускаем сервер заново через VBS-лаунчер
$Vbs = Join-Path $AppDir 'start_1c_bridge_silent.vbs'
if (-not (Test-Path $Vbs)) {
    Write-Host "ОШИБКА: не найден лаунчер $Vbs" -ForegroundColor Red
    Write-Host "RESTART_FAIL"
    if ($Interactive) { Read-Host "`nНажмите Enter для закрытия..." }
    exit 1
}
Write-Host "Запускаю $Vbs ..."
Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\wscript.exe') -ArgumentList "`"$Vbs`""

# 3) Ждём, пока сервер откроет порт 8000
$deadline = (Get-Date).AddSeconds(15)
$listener = $null
do {
    Start-Sleep -Milliseconds 500
    $listener = Get-NetTCPConnection -LocalPort 8000 -State Listen -ErrorAction SilentlyContinue |
        Select-Object -First 1
} while (-not $listener -and (Get-Date) -lt $deadline)

if (-not $listener) {
    Write-Host "ОШИБКА: сервер не поднялся. Смотрите %PROGRAMDATA%\1cMcpBridge\http-server.log" -ForegroundColor Red
    Write-Host "RESTART_FAIL"
    if ($Interactive) { Read-Host "`nНажмите Enter для закрытия..." }
    exit 1
}

# 4) Настоящий MCP health-check: tools/list через venv python
#    (токен передаём переменной окружения, не во временный файл)
$VenvPython = Join-Path $AppDir '.venv\Scripts\python.exe'
$TokenFile = Join-Path $env:ProgramData '1cMcpBridge\.http_token'
$token = ''
if (Test-Path $TokenFile) {
    $token = ([IO.File]::ReadAllText($TokenFile)).Trim()
}
if (-not $token) {
    Write-Host "ОШИБКА: не найден токен ($TokenFile)" -ForegroundColor Red
    Write-Host "RESTART_FAIL"
    if ($Interactive) { Read-Host "`nНажмите Enter для закрытия..." }
    exit 1
}

$env:ONEC_HTTP_TOKEN = $token
# Единый MCP health-check (src/healthcheck.py): ровно 5 инструментов +
# list_databases. Токен — только в переменной окружения.
$hcLine = $null
try {
    $hcLine = [string](& $VenvPython (Join-Path $AppDir 'healthcheck.py') 2>&1 | Select-Object -Last 1)
    if ($LASTEXITCODE -ne 0) { $hcLine = $null }
} finally {
    Remove-Item Env:\ONEC_HTTP_TOKEN -ErrorAction SilentlyContinue
}

if ($null -eq $hcLine -or $hcLine -notlike 'HEALTH_OK*') {
    Write-Host "ОШИБКА: MCP health-check не прошёл ($hcLine). Смотрите журнал сервера." -ForegroundColor Red
    Write-Host "RESTART_FAIL"
    if ($Interactive) { Read-Host "`nНажмите Enter для закрытия..." }
    exit 1
}

Write-Host "OK: HTTP-сервер работает (PID $($listener.OwningProcess), порт 8000, $hcLine)." -ForegroundColor Green
Write-Host "RESTART_OK"
if ($Interactive) { Read-Host "`nНажмите Enter для закрытия..." }
exit 0
