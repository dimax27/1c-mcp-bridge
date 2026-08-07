# =============================================================================
#  status_http_server.ps1 — статус HTTP-сервера 1C MCP Bridge.
#
#  Показывает: запущен ли процесс моста, слушает ли порт 8000, есть ли токен,
#  настроен ли ChatGPT/Codex (config.toml) и последние строки журнала.
#
#  Использование:
#    powershell -ExecutionPolicy Bypass -File status_http_server.ps1 [-Interactive]
# =============================================================================

[CmdletBinding()]
param(
    [switch]$Interactive  # при запуске из ярлыка: пауза после завершения
)

$ErrorActionPreference = 'Continue'

Write-Host "=== Статус 1C MCP Bridge (HTTP-сервер) ===" -ForegroundColor Cyan

# 1) Процесс моста (venv-python — лаунчер-редиректор, базовый python — рабочий)
$portOwner = Get-NetTCPConnection -LocalPort 8000 -State Listen -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty OwningProcess
$procs = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object {
        $_.ProcessId -ne $PID -and
        $_.CommandLine -and $_.CommandLine -match 'mcp_server_1c_http\.py'
    }
if ($procs) {
    foreach ($p in $procs) {
        $role = if ($p.ProcessId -eq $portOwner) { 'рабочий (слушает порт)' } else { 'лаунчер' }
        $exe = Split-Path $p.ExecutablePath -Leaf -ErrorAction SilentlyContinue
        Write-Host ("Процесс: PID {0} [{1}] ({2}, запущен {3})" -f $p.ProcessId, $role, $exe, $p.CreationDate)
    }
} else {
    Write-Host "Процесс: НЕ запущен" -ForegroundColor Yellow
    Write-Host "STATUS_PROCESS=NOT_RUNNING"
}

# 2) Порт 8000
$listener = Get-NetTCPConnection -LocalPort 8000 -State Listen -ErrorAction SilentlyContinue |
    Select-Object -First 1
if ($listener) {
    Write-Host "Порт 8000: слушает (PID $($listener.OwningProcess))" -ForegroundColor Green
    Write-Host "STATUS_PORT=LISTENING"
} else {
    Write-Host "Порт 8000: НЕ слушает" -ForegroundColor Red
    Write-Host "STATUS_PORT=NOT_LISTENING"
}

# 3) Токен
$tokenFile = Join-Path $env:ProgramData '1cMcpBridge\.http_token'
if (Test-Path $tokenFile) {
    $len = ([System.IO.File]::ReadAllText($tokenFile)).Trim().Length
    Write-Host "Токен: есть (длина $len)" -ForegroundColor Green
    Write-Host "STATUS_TOKEN=PRESENT"
} else {
    Write-Host "Токен: НЕ найден ($tokenFile)" -ForegroundColor Red
    Write-Host "STATUS_TOKEN=MISSING"
}

# 4) Конфиг ChatGPT/Codex
$cfg = Join-Path $env:USERPROFILE '.codex\config.toml'
if (Test-Path $cfg) {
    # config.toml в UTF-8 без BOM — читаем явно в UTF-8 (PS 5.1 по умолчанию ANSI)
    $text = Get-Content -LiteralPath $cfg -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
    if ($text -match '\[mcp_servers\.1c-bridge\]' -and $text -match '/mcp/') {
        Write-Host "ChatGPT/Codex config: секция 1c-bridge есть" -ForegroundColor Green
        Write-Host "STATUS_CODEX_CONFIG=PRESENT"
    } else {
        Write-Host "ChatGPT/Codex config: секции 1c-bridge нет" -ForegroundColor Yellow
        Write-Host "STATUS_CODEX_CONFIG=MISSING"
    }
} else {
    Write-Host "ChatGPT/Codex config: файл не найден ($cfg)" -ForegroundColor Yellow
    Write-Host "STATUS_CODEX_CONFIG=MISSING"
}

# 5) Последние строки журнала (журнал пишется в UTF-8 — читаем явно в UTF-8)
$logFile = Join-Path $env:ProgramData '1cMcpBridge\http-server.log'
if (Test-Path $logFile) {
    Write-Host "`n--- Последние строки http-server.log ---"
    Get-Content -LiteralPath $logFile -Tail 4 -Encoding UTF8 -ErrorAction SilentlyContinue |
        ForEach-Object {
            $_ -replace '/mcp/[A-Za-z0-9_-]{12,}', '/mcp/<token>' `
               -replace 'Pwd="[^"]*"', 'Pwd="***"'
        }
}

if ($Interactive) { Read-Host "`nНажмите Enter для закрытия..." }
exit 0
