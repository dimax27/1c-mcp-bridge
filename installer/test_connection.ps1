# =============================================================================
#  test_connection.ps1 — пробует подключиться к 1С с заданными параметрами.
#  Совместим с Windows PowerShell 5.1 (try/catch только как statement, не expr).
# =============================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$ParamsFile,
    [Parameter(Mandatory)] [string]$OutputFile
)

function Out([string]$line) {
    try {
        Add-Content -Path $OutputFile -Value $line -Encoding UTF8 -ErrorAction Stop
    } catch {
        [Console]::Error.WriteLine("FATAL: cannot write to ${OutputFile}: $_")
    }
}

# Создаём пустой файл сразу
try {
    Set-Content -Path $OutputFile -Value '' -Encoding UTF8 -ErrorAction Stop
} catch {
    [Console]::Error.WriteLine("FATAL: cannot create ${OutputFile}: $_")
    exit 10
}

Out "=== Тест подключения к 1С ==="
Out ("PowerShell version: " + $PSVersionTable.PSVersion)
Out ""

try {
    if (-not (Test-Path $ParamsFile)) {
        Out "ОШИБКА: Не найден файл параметров: $ParamsFile"
        exit 2
    }

    Out "Читаю параметры..."
    $params = @{}
    foreach ($line in Get-Content $ParamsFile -Encoding Default) {
        if ($line -match '^([^=]+)=(.*)$') {
            $params[$matches[1].Trim()] = $matches[2].Trim()
        }
    }

    $ProgID  = $params['PROGID']
    $ConnStr = $params['CONNSTR']
    $DllPath = $params['DLLPATH']

    if (-not $ProgID -or -not $ConnStr) {
        Out "ОШИБКА: пустые PROGID или CONNSTR."
        exit 3
    }

    Out "  ProgID:  $ProgID"
    $safeConn = $ConnStr -replace 'Pwd="[^"]*"', 'Pwd="***"'
    Out "  ConnStr: $safeConn"
    Out ""

    Out "Проверяю регистрацию COM-коннектора..."
    $type = [Type]::GetTypeFromProgID($ProgID, $false)
    if (-not $type) {
        Out "  $ProgID не зарегистрирован."
        if ($DllPath -and (Test-Path $DllPath)) {
            Out "  Регистрирую $DllPath..."
            $p = Start-Process -FilePath 'regsvr32.exe' `
                               -ArgumentList @('/s', "`"$DllPath`"") `
                               -Wait -PassThru -ErrorAction Stop
            if ($p.ExitCode -ne 0) {
                Out "  regsvr32 вернул код $($p.ExitCode). Нужны права администратора."
                exit 5
            }
            Out "  Зарегистрирован."
            $type = [Type]::GetTypeFromProgID($ProgID, $false)
            if (-not $type) {
                Out "  regsvr32 успешен, но ProgID не виден — несовпадение разрядности."
                exit 6
            }
        } else {
            Out "  comcntr.dll не найден ($DllPath)."
            Out "  Установщик попробует зарегистрировать его на этапе установки."
            Out "  Тест пропущен — можно нажимать Далее."
            exit 0
        }
    } else {
        Out "  OK — зарегистрирован."
    }
    Out ""

    Out "Создаю COM-объект..."
    $connector = [Activator]::CreateInstance($type)
    Out "  OK"

    Out "Подключаюсь к информационной базе..."
    $ib = $connector.Connect($ConnStr)
    Out "  OK"
    Out ""

    Out "Читаю метаданные..."
    $name = $null
    $synonym = $null
    try {
        $name = $ib.Метаданные.Имя
    } catch {
        Out "  Не удалось прочитать Метаданные.Имя: $($_.Exception.Message)"
    }
    try {
        $synonym = $ib.Метаданные.Синоним
    } catch {
        # некритично
    }

    Out ""
    Out "==============================================="
    Out "  УСПЕХ — подключение работает"
    Out "==============================================="
    if ($name)    { Out "Имя конфигурации: $name" }
    if ($synonym) { Out "Синоним:          $synonym" }
    Out ""
    Out "Можно нажимать Далее."
    exit 0

} catch {
    Out ""
    Out "==============================================="
    Out "  ОШИБКА"
    Out "==============================================="
    Out $_.Exception.Message
    if ($_.Exception.InnerException) {
        Out ""
        Out "Inner: $($_.Exception.InnerException.Message)"
    }
    Out ""
    Out "Возможные причины:"
    Out "  - Неверный адрес сервера или имя ИБ"
    Out "  - Нет свободной клиентской лицензии 1С"
    Out "  - У учётки нет прав на эту базу"
    Out "  - Брандмауэр блокирует порт сервера 1С (1541)"
    Out "  - Несовпадение версии COM-коннектора и сервера"
    exit 7
}
