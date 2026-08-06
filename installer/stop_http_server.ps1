# =============================================================================
#  stop_http_server.ps1 — остановка запущенного HTTP-сервера 1C MCP Bridge.
#
#  Зачем: при переустановке старый сервер держит порт 8000 и файлы venv —
#  новый установщик не сможет занять порт, а удаление .venv упрётся в
#  блокировку. Этот скрипт находит и останавливает старый сервер ДО начала
#  установки/удаления.
#
#  Использование:
#    напрямую:   powershell -ExecutionPolicy Bypass -File stop_http_server.ps1
#    из скрипта: . (Join-Path $PSScriptRoot 'stop_http_server.ps1')
#                Stop-BridgeHttpServer
# =============================================================================

[CmdletBinding()]
param()

function Stop-BridgeHttpServer {
    [CmdletBinding()]
    param(
        [int]$Port = 8000,
        [int]$WaitSeconds = 10
    )

    $stopped = @()

    # 1) Процессы моста по командной строке (надёжнее, чем путь к exe)
    try {
        Get-CimInstance Win32_Process -ErrorAction Stop |
            Where-Object { $_.CommandLine -and $_.CommandLine -match 'mcp_server_1c_http\.py' } |
            ForEach-Object {
                $stopped += $_
                Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
                Write-Host "Останавливаю HTTP-сервер моста: PID $($_.ProcessId)"
                Write-Host "BRIDGE_HTTP_STOPPED: $($_.ProcessId)"
            }
    } catch {
        Write-Warning "Не удалось получить список процессов: $($_.Exception.Message)"
    }

    # 2) Владелец порта 8000 — на случай, если CommandLine недоступен
    try {
        Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue |
            ForEach-Object {
                $owner = Get-CimInstance Win32_Process `
                    -Filter "ProcessId = $($_.OwningProcess)" `
                    -ErrorAction SilentlyContinue
                $isBridge = $owner -and $owner.CommandLine -and `
                    $owner.CommandLine -match 'mcp_server_1c_http\.py'
                $alreadyStopped = $stopped | ForEach-Object { $_.ProcessId }
                if ($isBridge -and ($owner.ProcessId -notin $alreadyStopped)) {
                    Stop-Process -Id $owner.ProcessId -Force -ErrorAction SilentlyContinue
                    Write-Host "Останавливаю владельца порта ${Port}: PID $($owner.ProcessId)"
                }
            }
    } catch {
        # Get-NetTCPConnection может отсутствовать в старых ОС — не критично
    }

    # 3) Ждём освобождения порта
    $deadline = (Get-Date).AddSeconds($WaitSeconds)
    do {
        $listener = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
        if (-not $listener) { break }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)

    if (Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue) {
        Write-Warning "Порт $Port всё ещё занят после остановки моста."
        Write-Host "PORT_${Port}_BUSY"
        return $false
    }

    Write-Host "Порт $Port свободен."
    Write-Host "PORT_${Port}_FREE"
    return $true
}

# Прямой запуск (не dot-source): выполняем и завершаемся кодом результата
if ($MyInvocation.InvocationName -ne '.') {
    $ok = Stop-BridgeHttpServer
    exit $(if ($ok) { 0 } else { 1 })
}
