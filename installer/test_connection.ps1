# =============================================================================
#  test_connection.ps1 — пробует подключиться к 1С с заданными параметрами.
#  Вызывается из Pascal Script кнопкой "Проверить подключение к 1С".
#  Не зависит от установленного venv: использует системный Python или сразу COM.
# =============================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$ParamsFile,
    [Parameter(Mandatory)] [string]$OutputFile
)

$ErrorActionPreference = 'Stop'

function Out([string]$line) {
    Add-Content -Path $OutputFile -Value $line -Encoding UTF8
}

# Очистим output
Set-Content -Path $OutputFile -Value '' -Encoding UTF8

try {
    if (-not (Test-Path $ParamsFile)) {
        Out "Не найден файл параметров: $ParamsFile"
        exit 2
    }

    $params = @{}
    foreach ($line in Get-Content $ParamsFile -Encoding UTF8) {
        if ($line -match '^([^=]+)=(.*)$') {
            $params[$matches[1].Trim()] = $matches[2].Trim()
        }
    }

    $ProgID  = $params['PROGID']
    $ConnStr = $params['CONNSTR']
    $DllPath = $params['DLLPATH']

    Out "Параметры:"
    Out "  ProgID:  $ProgID"
    $safeConn = $ConnStr -replace 'Pwd="[^"]*"', 'Pwd="***"'
    Out "  ConnStr: $safeConn"
    Out ""

    # Если коннектор не зарегистрирован — попробуем зарегистрировать прямо сейчас
    # (для теста до полной установки)
    try {
        $type = [Type]::GetTypeFromProgID($ProgID, $false)
        if (-not $type) {
            Out "ProgID '$ProgID' не зарегистрирован в системе."
            if ($DllPath -and (Test-Path $DllPath)) {
                Out "Регистрирую $DllPath ..."
                $p = Start-Process -FilePath 'regsvr32.exe' -ArgumentList @('/s', "`"$DllPath`"") -Wait -PassThru
                if ($p.ExitCode -ne 0) {
                    Out "regsvr32 вернул код $($p.ExitCode)."
                    exit 3
                }
                Out "Зарегистрирован."
            } else {
                Out "comcntr.dll не найден по пути $DllPath."
                Out "Регистрация невозможна. Завершаю проверку с ошибкой."
                exit 4
            }
        }

        Out "Создаю COM-объект $ProgID ..."
        $connector = [Activator]::CreateInstance([Type]::GetTypeFromProgID($ProgID))

        Out "Подключаюсь к информационной базе ..."
        $ib = $connector.Connect($ConnStr)

        $name    = $ib.Метаданные.Имя
        $synonym = $ib.Метаданные.Синоним

        Out ""
        Out "УСПЕХ"
        Out "Имя конфигурации: $name"
        if ($synonym) { Out "Синоним:          $synonym" }
        Out ""
        Out "Подключение работает. Можно нажимать «Далее»."
        exit 0

    } catch {
        Out ""
        Out "ОШИБКА ПОДКЛЮЧЕНИЯ:"
        Out $_.Exception.Message
        if ($_.Exception.InnerException) {
            Out ""
            Out "Inner: $($_.Exception.InnerException.Message)"
        }
        Out ""
        Out "Возможные причины:"
        Out "  • Неверный адрес сервера или имя ИБ"
        Out "  • Нет свободной лицензии 1С"
        Out "  • Учётка не имеет прав на эту базу"
        Out "  • Брандмауэр блокирует порт сервера 1С"
        exit 5
    }

} catch {
    Out "Внутренняя ошибка теста: $($_.Exception.Message)"
    exit 99
}
