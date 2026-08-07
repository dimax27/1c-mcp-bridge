# =============================================================================
#  stop_http_server.ps1 — остановка запущенного HTTP-сервера 1C MCP Bridge.
#
#  Зачем: при переустановке старый сервер держит порт 8000 и файлы venv —
#  новый установщик не сможет занять порт, а удаление .venv упрётся в
#  блокировку. Этот скрипт находит и останавливает старый сервер ДО начала
#  установки/удаления.
#
#  Безопасность остановки: по умолчанию (без параметров) останавливаются
#  ВСЕ процессы с mcp_server_1c_http.py ИЛИ mcp_server_1c.py в командной строке — для удаления
#  это правильно. При установке передавайте ExpectedScriptPath /
#  ExpectedPythonPath: тогда остановится ТОЛЬКО процесс этой установки,
#  а посторонние экземпляры (тестовые, другой каталог) не пострадают.
#
#  Использование:
#    напрямую:   powershell -ExecutionPolicy Bypass -File stop_http_server.ps1
#    из скрипта: . (Join-Path $PSScriptRoot 'stop_http_server.ps1')
#                Stop-BridgeHttpServer -Port 8000 `
#                    -ExpectedScriptPath "$AppDir\mcp_server_1c_http.py" `
#                    -ExpectedPythonPath "$AppDir\.venv\Scripts\python.exe"
# =============================================================================

[CmdletBinding()]
param(
    [switch]$Interactive,  # при запуске из ярлыка: пауза после завершения
    [int]$Port = 8000,
    [int]$WaitSeconds = 10,
    [string]$ExpectedScriptPath = '',
    [string]$ExpectedPythonPath = ''
)

function Stop-BridgeHttpServer {
    [CmdletBinding()]
    param(
        [int]$Port = 8000,
        [int]$WaitSeconds = 10,
        [string]$ExpectedScriptPath = '',
        [string]$ExpectedPythonPath = ''
    )

    $stopped = @()

    # 1) Процессы моста по командной строке (надёжнее, чем путь к exe).
    #    При заданных Expected* путях — только процесс этой установки.
    try {
        # $PID исключаем: командная строка скрипта сама содержит
        # -ExpectedScriptPath .../mcp_server_1c_http.py и ловится фильтром.
        $bridgeProcs = Get-CimInstance Win32_Process -ErrorAction Stop |
            Where-Object {
                $_.ProcessId -ne $PID -and
                $_.CommandLine -and $_.CommandLine -match 'mcp_server_1c(?:_http)?\.py'
            }

        if ($ExpectedScriptPath) {
            $bridgeProcs = $bridgeProcs |
                Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -like "*$ExpectedScriptPath*" }
        }
        if ($ExpectedPythonPath) {
            $bridgeProcs = $bridgeProcs |
                Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -like "*$ExpectedPythonPath*" }
        }

        $bridgeProcs | ForEach-Object {
            $stopped += $_
            Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
            Write-Host "Останавливаю HTTP-сервер моста: PID $($_.ProcessId)"
            Write-Host "BRIDGE_HTTP_STOPPED: $($_.ProcessId)"
        }
    } catch {
        Write-Warning "Не удалось получить список процессов: $($_.Exception.Message)"
    }

    # 2) Владелец порта $Port — на случай, если CommandLine-фильтр не сработал.
    #    При заданных Expected* путях владелец останавливается, только если он
    #    принадлежит этой установке; посторонний процесс на порту не трогаем
    #    (тогда функция вернёт $false и установка прервётся).
    try {
        Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue |
            ForEach-Object {
                $owner = Get-CimInstance Win32_Process `
                    -Filter "ProcessId = $($_.OwningProcess)" `
                    -ErrorAction SilentlyContinue
                if (-not $owner) { return }
                $isBridge = $owner.CommandLine -and `
                    $owner.CommandLine -match 'mcp_server_1c(?:_http)?\.py' -and `
                    $owner.ProcessId -ne $PID
                if ($ExpectedScriptPath) {
                    $isBridge = $isBridge -and `
                        $owner.CommandLine -like "*$ExpectedScriptPath*"
                }
                if ($ExpectedPythonPath) {
                    $isBridge = $isBridge -and `
                        $owner.CommandLine -like "*$ExpectedPythonPath*"
                }
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
    $ok = Stop-BridgeHttpServer -Port $Port -WaitSeconds $WaitSeconds `
        -ExpectedScriptPath $ExpectedScriptPath `
        -ExpectedPythonPath $ExpectedPythonPath
    if ($Interactive) {
        Read-Host "`nНажмите Enter для закрытия..."
    }
    exit $(if ($ok) { 0 } else { 1 })
}
