# =============================================================================
#  restart_http_server.ps1 — перезапуск HTTP-сервера 1C MCP Bridge.
#
#  Останавливает старый сервер (stop_http_server.ps1), запускает VBS-лаунчер
#  заново и проверяет, что порт 8000 слушает.
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

# 1) Останавливаем старый сервер
$StopScript = Join-Path $PSScriptRoot 'stop_http_server.ps1'
if (-not (Test-Path $StopScript)) {
    Write-Host "ОШИБКА: не найден $StopScript" -ForegroundColor Red
    if ($Interactive) { Read-Host "`nНажмите Enter для закрытия..." }
    exit 1
}
. $StopScript
$AppDir = Split-Path $PSScriptRoot -Parent
$stopped = Stop-BridgeHttpServer `
    -Port 8000 `
    -ExpectedScriptPath (Join-Path $AppDir 'mcp_server_1c_http.py') `
    -ExpectedPythonPath (Join-Path $AppDir '.venv\Scripts\python.exe')

# 2) Запускаем сервер заново через VBS-лаунчер
$Vbs = Join-Path (Split-Path $PSScriptRoot -Parent) 'start_1c_bridge_silent.vbs'
if (-not (Test-Path $Vbs)) {
    Write-Host "ОШИБКА: не найден лаунчер $Vbs" -ForegroundColor Red
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

if ($listener) {
    Write-Host "OK: HTTP-сервер работает (PID $($listener.OwningProcess), порт 8000)." -ForegroundColor Green
    Write-Host "RESTART_OK"
    if ($Interactive) { Read-Host "`nНажмите Enter для закрытия..." }
    exit 0
}

Write-Host "ОШИБКА: сервер не поднялся. Смотрите %PROGRAMDATA%\1cMcpBridge\http-server.log" -ForegroundColor Red
Write-Host "RESTART_FAIL"
if ($Interactive) { Read-Host "`nНажмите Enter для закрытия..." }
exit 1
