# =============================================================================
#  install.ps1 — основная установка после копирования файлов мастером.
#  Параметры читаются из install_params.txt, сохранённого Pascal Script'ом.
#
#  Шаги:
#   1. Чтение параметров.
#   2. Поиск/установка Python 3.12 (если нет system-wide).
#   3. Создание venv в %APP%\.venv.
#   4. Установка зависимостей.
#   5. regsvr32 для COM-коннектора выбранной платформы 1С.
#   6. Запись блока 1c-bridge в claude_desktop_config.json (UTF-8 без BOM).
# =============================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# Лог пишем рядом с инсталлером — потом удобно дебажить
$LogPath = Join-Path $PSScriptRoot 'install.log'
function Log {
    param([string]$Message)
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$ts  $Message" | Tee-Object -FilePath $LogPath -Append | Out-Host
}

trap {
    Log ("ОШИБКА: " + $_.Exception.Message)
    Log ($_.ScriptStackTrace)
    exit 1
}

Log "Запуск install.ps1"

# -----------------------------------------------------------------------------
# 1. Параметры
# -----------------------------------------------------------------------------
$ParamsFile = Join-Path $PSScriptRoot 'install_params.txt'
if (-not (Test-Path $ParamsFile)) {
    throw "Не найден файл параметров $ParamsFile"
}

$params = @{}
foreach ($line in Get-Content $ParamsFile -Encoding UTF8) {
    if ($line -match '^([^=]+)=(.*)$') {
        $params[$matches[1].Trim()] = $matches[2].Trim()
    }
}

$ProgID      = $params['PROGID']
$ConnStr     = $params['CONNSTR']
$DllPath     = $params['DLLPATH']
$AppDir      = $params['APPDIR']
$UserAppData = $params['USERAPPDATA']

Log "ProgID       = $ProgID"
Log "ConnStr      = $($ConnStr -replace 'Pwd="[^"]*"', 'Pwd="***"')"
Log "DllPath      = $DllPath"
Log "AppDir       = $AppDir"
Log "UserAppData  = $UserAppData"

# -----------------------------------------------------------------------------
# 2. Python 3.12 — ищем готовый, иначе ставим
# -----------------------------------------------------------------------------
function Find-Python312 {
    # Пробуем py-launcher
    try {
        $v = & py -3.12 --version 2>&1
        if ($LASTEXITCODE -eq 0 -and $v -match 'Python 3\.12') {
            $exe = & py -3.12 -c "import sys; print(sys.executable)" 2>&1
            if ($LASTEXITCODE -eq 0) { return $exe.Trim() }
        }
    } catch { }

    # Прямой поиск в стандартных папках
    $candidates = @(
        "$env:ProgramFiles\Python312\python.exe",
        "${env:ProgramFiles(x86)}\Python312\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe"
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { return $c }
    }
    return $null
}

$PythonExe = Find-Python312
if (-not $PythonExe) {
    Log "Python 3.12 не найден. Скачиваю и устанавливаю..."
    $url = 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe'
    $tmp = Join-Path $env:TEMP 'python-3.12.7-amd64.exe'

    Log "Загрузка $url"
    Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing

    Log "Запуск тихой установки Python (для всех пользователей, в PATH)..."
    $proc = Start-Process -FilePath $tmp -Wait -PassThru -ArgumentList @(
        '/quiet',
        'InstallAllUsers=1',
        'PrependPath=1',
        'Include_launcher=1',
        'Include_test=0'
    )
    if ($proc.ExitCode -ne 0) {
        throw "Установка Python завершилась с кодом $($proc.ExitCode)"
    }

    $PythonExe = Find-Python312
    if (-not $PythonExe) {
        throw "Python установлен, но не нашёл python.exe. Проверь $env:ProgramFiles\Python312\."
    }
}
Log "Python: $PythonExe"

# -----------------------------------------------------------------------------
# 3. venv
# -----------------------------------------------------------------------------
$VenvDir = Join-Path $AppDir '.venv'
if (Test-Path $VenvDir) {
    Log "Удаляю старый venv..."
    Remove-Item -Recurse -Force $VenvDir
}
Log "Создаю venv в $VenvDir"
& $PythonExe -m venv $VenvDir
if ($LASTEXITCODE -ne 0) { throw "python -m venv упал, код $LASTEXITCODE" }

$VenvPython = Join-Path $VenvDir 'Scripts\python.exe'

# -----------------------------------------------------------------------------
# 4. Зависимости
# -----------------------------------------------------------------------------
Log "Обновляю pip..."
& $VenvPython -m pip install --upgrade pip 2>&1 | Tee-Object -FilePath $LogPath -Append

Log "Устанавливаю зависимости из requirements.txt..."
& $VenvPython -m pip install -r (Join-Path $AppDir 'requirements.txt') 2>&1 |
    Tee-Object -FilePath $LogPath -Append
if ($LASTEXITCODE -ne 0) { throw "pip install вернул $LASTEXITCODE" }

# -----------------------------------------------------------------------------
# 5. Регистрация COM-коннектора
# -----------------------------------------------------------------------------
if ($DllPath -and (Test-Path $DllPath)) {
    Log "Регистрирую $DllPath..."
    $proc = Start-Process -FilePath 'regsvr32.exe' -ArgumentList @('/s', "`"$DllPath`"") -Wait -PassThru
    if ($proc.ExitCode -ne 0) {
        Log "regsvr32 вернул код $($proc.ExitCode) — возможно, коннектор уже зарегистрирован."
    } else {
        Log "COM-коннектор зарегистрирован."
    }

    # Массовая регистрация остальных DLL из bin'а 1С.
    # При первом подключении через V83.COMConnector платформа подгружает
    # type-libraries из соседних DLL (frnt*, bsl*, wbas*, …). Если они не
    # зарегистрированы — Connect() падает с TYPE_E_LIBNOTREGISTERED (0x8002801D).
    # Регистрируем все доступные DLL — лишнего не будет, regsvr32 для не-COM
    # библиотек просто молча пропустит.
    $binDir = Split-Path $DllPath -Parent
    if (Test-Path $binDir) {
        Log "Регистрирую остальные DLL из $binDir (для type-libraries)..."
        $count = 0
        Get-ChildItem $binDir -Filter '*.dll' -ErrorAction SilentlyContinue | ForEach-Object {
            # comcntr уже сделан выше — пропускаем
            if ($_.Name -ieq 'comcntr.dll') { return }
            # /s - silent, /i - вызов DllInstall если есть
            Start-Process -FilePath 'regsvr32.exe' `
                          -ArgumentList @('/s', "`"$($_.FullName)`"") `
                          -Wait -PassThru -ErrorAction SilentlyContinue | Out-Null
            $count++
        }
        Log "Обработано DLL: $count (часть из них не COM — это нормально)."
    }
} else {
    Log "Путь к comcntr.dll не задан или не существует ($DllPath). Пропускаю regsvr32."
    Log "Если потом возникнет ошибка 'Class not registered' — выполни вручную: regsvr32 <путь к comcntr.dll>"
}

# -----------------------------------------------------------------------------
# 6. databases.json — генерируем на основе параметров мастера
# -----------------------------------------------------------------------------
$DatabasesFile = Join-Path $AppDir 'databases.json'
$DbKey  = $params['DBKEY']
$DbDesc = $params['DBDESC']
if (-not $DbKey)  { $DbKey  = 'main' }

# Если databases.json уже существует (переустановка) — добавляем/обновляем нашу базу
if (Test-Path $DatabasesFile) {
    Log "Найден существующий $DatabasesFile — обновляю запись '$DbKey'."
    try {
        $dbConfig = Get-Content $DatabasesFile -Raw -Encoding UTF8 | ConvertFrom-Json -AsHashtable
        if (-not $dbConfig.databases) { $dbConfig.databases = @{} }
    } catch {
        Log "Не удалось прочитать databases.json: $($_.Exception.Message). Создаю заново."
        $dbConfig = @{ version = 1; default_database = ''; databases = @{} }
    }
} else {
    $dbConfig = @{ version = 1; default_database = ''; databases = @{} }
}

if ($dbConfig.databases -isnot [hashtable]) {
    $tmp = @{}
    foreach ($p in $dbConfig.databases.PSObject.Properties) { $tmp[$p.Name] = $p.Value }
    $dbConfig.databases = $tmp
}

$dbEntry = @{
    description = if ($DbDesc) { $DbDesc } else { $DbKey }
    progid = $ProgID
    connection_string = $ConnStr
    notes = ''
}
if ($DllPath) { $dbEntry.dll_path = $DllPath }

$dbConfig.databases[$DbKey] = $dbEntry
if (-not $dbConfig.default_database -or $dbConfig.databases.Keys.Count -eq 1) {
    $dbConfig.default_database = $DbKey
}
$dbConfig.version = 1

$dbJson = $dbConfig | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText($DatabasesFile, $dbJson, [System.Text.UTF8Encoding]::new($false))
Log "Записан $DatabasesFile (база '$DbKey')"

# -----------------------------------------------------------------------------
# 7. claude_desktop_config.json
# -----------------------------------------------------------------------------
$ClaudeDir    = Join-Path $UserAppData 'Claude'
$ConfigPath   = Join-Path $ClaudeDir 'claude_desktop_config.json'

if (-not (Test-Path $ClaudeDir)) {
    Log "Папка $ClaudeDir не найдена — Claude Desktop ещё не запускался ни разу."
    Log "Создаю папку и кладу конфиг — Claude подхватит при первом запуске."
    New-Item -ItemType Directory -Path $ClaudeDir -Force | Out-Null
}

if (Test-Path $ConfigPath) {
    Log "Найден существующий config — обновляю блок 1c-bridge."
    try {
        $jsonText = Get-Content -Path $ConfigPath -Raw -Encoding UTF8
        $config   = $jsonText | ConvertFrom-Json -AsHashtable
    } catch {
        Log "Не удалось распарсить существующий config: $($_.Exception.Message)"
        Log "Делаю backup и создаю новый."
        Copy-Item $ConfigPath ($ConfigPath + '.bak.' + (Get-Date -Format 'yyyyMMddHHmmss')) -Force
        $config = @{}
    }
} else {
    $config = @{}
}

if (-not $config.ContainsKey('mcpServers')) {
    $config['mcpServers'] = @{}
}
if ($config.mcpServers -isnot [hashtable]) {
    $tmp = @{}
    foreach ($p in $config.mcpServers.PSObject.Properties) { $tmp[$p.Name] = $p.Value }
    $config.mcpServers = $tmp
}

# В v0.2.0 — указываем путь к databases.json через переменную окружения
$config.mcpServers['1c-bridge'] = @{
    command = $VenvPython
    args    = @( (Join-Path $AppDir 'mcp_server_1c.py') )
    env     = @{
        ONEC_DATABASES_FILE = $DatabasesFile
    }
}

$json = $config | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText($ConfigPath, $json, [System.Text.UTF8Encoding]::new($false))
Log "Конфиг обновлён: $ConfigPath"

Log "Установка завершена успешно."
exit 0
