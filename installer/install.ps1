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

# Крупный заголовок этапа — чтобы пользователь видел что сейчас делается
$script:StageNum = 0
function Stage {
    param([string]$Title)
    $script:StageNum++
    $bar = "=" * 70
    Write-Host ""
    Write-Host $bar -ForegroundColor Cyan
    Write-Host (" Этап {0} : {1}" -f $script:StageNum, $Title) -ForegroundColor Cyan
    Write-Host $bar -ForegroundColor Cyan
    Log "[Этап $script:StageNum] $Title"
}

trap {
    Log ("ОШИБКА: " + $_.Exception.Message)
    Log ($_.ScriptStackTrace)
    Write-Host ""
    Write-Host "ОШИБКА: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Окно закроется через 30 секунд. Можно сфотографировать ошибку." -ForegroundColor Yellow
    Start-Sleep -Seconds 30
    exit 1
}

Log "Запуск install.ps1"

# -----------------------------------------------------------------------------
Stage "Чтение параметров мастера"
# -----------------------------------------------------------------------------
$ParamsFile = Join-Path $PSScriptRoot 'install_params.txt'
if (-not (Test-Path $ParamsFile)) {
    throw "Не найден файл параметров $ParamsFile"
}

$params = @{}
try {
    foreach ($line in Get-Content -LiteralPath $ParamsFile -Encoding Default) {
        if ($line -match '^([^=]+)=(.*)$') {
            $params[$matches[1].Trim()] = $matches[2].Trim()
        }
    }
}
finally {
    if (Test-Path -LiteralPath $ParamsFile) {
        Remove-Item -LiteralPath $ParamsFile -Force -ErrorAction Stop
    }
}

$ProgID      = $params['PROGID']
$ConnStr     = $params['CONNSTR']
$DllPath     = $params['DLLPATH']
$AppDir      = $params['APPDIR']
$UserAppData = $params['USERAPPDATA']
$ImportDbFile = $params['IMPORT_DB_FILE']

Log "ProgID       = $ProgID"
Log "ConnStr      = $($ConnStr -replace 'Pwd="[^"]*"', 'Pwd="***"')"
Log "DllPath      = $DllPath"
Log "AppDir       = $AppDir"
Log "UserAppData  = $UserAppData"

# -----------------------------------------------------------------------------
# 2. Python
# -----------------------------------------------------------------------------
Stage "Поиск/установка Python 3.12"
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
Stage "Создание изолированной Python-среды (venv)"
$VenvDir = Join-Path $AppDir '.venv'
if (Test-Path $VenvDir) {
    Log "Удаляю старый venv..."
    # Сначала убиваем все python из старого venv (Claude Desktop мог держать процесс)
    Get-Process python, pythonw -ErrorAction SilentlyContinue | Where-Object {
        try { $_.Path -and ($_.Path -like "$VenvDir*") } catch { $false }
    } | ForEach-Object {
        Log "Останавливаю процесс $($_.Id) ($($_.Path))..."
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2

    $attempts = 0
    while ((Test-Path $VenvDir) -and $attempts -lt 5) {
        try {
            Remove-Item -Recurse -Force $VenvDir -ErrorAction Stop
            break
        } catch {
            $attempts++
            Log "Попытка $attempts из 5: файл занят, жду 3 секунды..."
            Start-Sleep -Seconds 3
        }
    }
    if (Test-Path $VenvDir) {
        Log "Не удалось удалить $VenvDir. Закройте Claude Desktop полностью (Quit из трея) и попробуйте снова."
        throw "venv заблокирован: $VenvDir"
    }
}
Log "Создаю venv в $VenvDir"
& $PythonExe -m venv $VenvDir
if ($LASTEXITCODE -ne 0) { throw "python -m venv упал, код $LASTEXITCODE" }

$VenvPython = Join-Path $VenvDir 'Scripts\python.exe'

# -----------------------------------------------------------------------------
# 4. Зависимости
# -----------------------------------------------------------------------------
Stage "Установка Python-зависимостей (pywin32, mcp)"
Log "Обновляю pip..."
& $VenvPython -m pip install --upgrade pip 2>&1 | Tee-Object -FilePath $LogPath -Append

Log "Устанавливаю зависимости из requirements.txt..."
& $VenvPython -m pip install -r (Join-Path $AppDir 'requirements.txt') 2>&1 |
    Tee-Object -FilePath $LogPath -Append
if ($LASTEXITCODE -ne 0) { throw "pip install вернул $LASTEXITCODE" }

# -----------------------------------------------------------------------------
# 5. Регистрация COM-коннектора
# -----------------------------------------------------------------------------
Stage "Регистрация COM-коннектора 1С (это может занять 1-2 минуты)"

# Try using connector directly first — often already registered
$needRegistration = $true
try {
    $testConnector = New-Object -ComObject $ProgID
    Log "COM-коннектор уже зарегистрирован — пропускаю regsvr32."
    $needRegistration = $false
} catch {
    Log "COM-коннектор не зарегистрирован."
}

if ($needRegistration -and $DllPath -and (Test-Path $DllPath)) {
    Log "Регистрирую $DllPath..."
    $proc = Start-Process -FilePath 'regsvr32.exe' -ArgumentList @('/s', "`"$DllPath`"") -Wait -PassThru
    if ($proc.ExitCode -ne 0) { Log "regsvr32 exit code: $($proc.ExitCode)" }
    else { Log "COM-коннектор зарегистрирован." }

    # Register essential type-library DLLs only
    $binDir = Split-Path $DllPath -Parent
    $essential = @('comcntr.dll', 'comcntr64.dll', 'V8Reader.dll', 'V8Writer.dll',
                   'core82.dll', 'core83.dll', 'core85.dll', 'backend.dll')
    $toRegister = @()
    foreach ($dll in $essential) {
        $p = Join-Path $binDir $dll
        if (Test-Path $p) { $toRegister += $p }
    }
    if ($toRegister.Count -eq 0) {
        Log "Essential DLLs not found, falling back to bulk..."
        $toRegister = @(Get-ChildItem $binDir -Filter '*.dll' | % { $_.FullName })
    }
    Log "Registering $($toRegister.Count) DLLs for type-libraries..."
    $ok, $fail = 0, 0
    foreach ($dll in $toRegister) {
        $p = Start-Process -FilePath 'regsvr32.exe' -ArgumentList @('/s', "`"$dll`"") -Wait -PassThru
        if ($p.ExitCode -eq 0) { $ok++ } else { $fail++ }
    }
    Log "DLLs: $ok OK, $fail failed (non-COM DLLs fail normally)"
} elseif (-not $DllPath) {
    Log "comcntr.dll path not set — skipping registration."
}

# -----------------------------------------------------------------------------
# 6. databases.json — генерируем на основе параметров мастера
# -----------------------------------------------------------------------------
Stage "Создание databases.json"

# В v0.2.0+ файл лежит в ProgramData (доступен на запись обычным пользователям)
$DataDir = Join-Path $env:PROGRAMDATA '1cMcpBridge'
if (-not (Test-Path $DataDir)) {
    New-Item -ItemType Directory -Path $DataDir -Force | Out-Null
}
$DatabasesFile = Join-Path $DataDir 'databases.json'

# ---- ACL helper: set clean permissions on databases.json ----
function Set-DatabaseFileAcl {
    param($FilePath)
    # Extract interactive username from AppData path
    $userName = ($UserAppData -split '\\AppData\\')[0] -replace '.*\\', ''
    if (-not $userName) { Log "WARNING: cannot determine user from $UserAppData"; return }
    try {
        $userSid = (New-Object System.Security.Principal.NTAccount($userName)).Translate([System.Security.Principal.SecurityIdentifier])
        $acl = New-Object System.Security.AccessControl.FileSecurity
        $acl.SetAccessRuleProtection($true, $false)
        $acl.AddAccessRule([System.Security.AccessControl.FileSystemAccessRule]::new('S-1-5-18', 'FullControl', 'Allow'))
        $acl.AddAccessRule([System.Security.AccessControl.FileSystemAccessRule]::new('S-1-5-32-544', 'FullControl', 'Allow'))
        $acl.AddAccessRule([System.Security.AccessControl.FileSystemAccessRule]::new($userSid, 'Modify', 'Allow'))
        Set-Acl -LiteralPath $FilePath -AclObject $acl
        Log "ACL: SYSTEM+Admins(Full), $userName(Modify), inheritance off"
    } catch {
        Log "ACL .NET failed, trying icacls..."
        if (-not $userSid) {
            throw "Cannot determine SID for user '$userName'"
        }
        & icacls.exe $FilePath /inheritance:r /grant:r "*S-1-5-18:(F)" "*S-1-5-32-544:(F)" "*$($userSid.Value):(M)" 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to secure databases.json: icacls exit code $LASTEXITCODE"
        }
        Log "ACL via icacls for $userName"
    }
}

# --- Импорт существующего databases.json (из мастера установки) ---
if ($ImportDbFile -and (Test-Path $ImportDbFile)) {
    Log "Импортирую databases.json из $ImportDbFile ..."
    try {
        $importedJson = Get-Content $ImportDbFile -Raw -Encoding UTF8 | ConvertFrom-Json
        # Валидация: должен быть объект с ключом databases
        if (-not $importedJson.databases) {
            throw "Файл не содержит ключа 'databases'"
        }
        Copy-Item $ImportDbFile $DatabasesFile -Force
        Log "Импортирован $DatabasesFile из $ImportDbFile"

        # Set clean ACL for interactive user
        Set-DatabaseFileAcl $DatabasesFile
    } catch {
        Log "Ошибка импорта: $($_.Exception.Message). Создаю новую базу вручную."
        $ImportDbFile = $null  # сбрасываем, чтобы сработал ручной ввод ниже
    }
}

# Ручное создание (если не было импорта)
if (-not $ImportDbFile -or -not (Test-Path $ImportDbFile)) {

# Миграция со старого пути (v0.2.0-beta.1 и ранее)
$LegacyFile = Join-Path $AppDir 'databases.json'
if ((Test-Path $LegacyFile) -and -not (Test-Path $DatabasesFile)) {
    Log "Переношу $LegacyFile -> $DatabasesFile"
    Copy-Item $LegacyFile $DatabasesFile -Force
    # Старый файл удалим в самом конце, после успешной записи нового
}

$DbKey  = $params['DBKEY']
$DbDesc = $params['DBDESC']
if (-not $DbKey)  { $DbKey  = 'main' }

if (Test-Path $DatabasesFile) {
    Log "Найден существующий $DatabasesFile — обновляю запись '$DbKey'."
    try {
        $dbConfig = Get-Content $DatabasesFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $dbConfig = ConvertTo-HashtableDeep $dbConfig
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

# Set clean ACL for interactive user
Set-DatabaseFileAcl $DatabasesFile

# Удаляем legacy-файл (если был)
if ((Test-Path $LegacyFile) -and ($LegacyFile -ne $DatabasesFile)) {
    Remove-Item $LegacyFile -Force -ErrorAction SilentlyContinue
    Log "Удалён старый $LegacyFile"
}

}  # конец блока ручного создания (если не было импорта)

# -----------------------------------------------------------------------------
# 7. Configure MCP clients (Claude, Qwen, Kimi, Reasonix)
# -----------------------------------------------------------------------------
Stage "Configuring MCP clients"

# --- Setup npm launcher for Qwen (no spaces in path) ---
$NpxDir = Join-Path $env:PROGRAMDATA '1cMcpBridge'
$NpxPackageJson = Join-Path $NpxDir 'package.json'
$NpxLauncherCmd = Join-Path $NpxDir 'launcher.cmd'

# Write package.json for npx
$packageJson = @"
{
    "name": "1c-mcp-bridge-launcher",
    "version": "1.0.0",
    "description": "Launch 1C MCP Bridge Python server from npx",
    "bin": { "1c-bridge": "./launcher.cmd" },
    "private": true
}
"@
[System.IO.File]::WriteAllText($NpxPackageJson, $packageJson, [System.Text.UTF8Encoding]::new($false))

# Write launcher.cmd with actual paths
$launcherCmd = @"
@echo off
setlocal
if not defined ONEC_DATABASES_FILE set ONEC_DATABASES_FILE=$DatabasesFile
"$VenvPython" "$(Join-Path $AppDir 'mcp_server_1c.py')"
"@
[System.IO.File]::WriteAllText($NpxLauncherCmd, $launcherCmd, [System.Text.UTF8Encoding]::new($false))
Log "Created npx launcher in $NpxDir"

# --- Create silent VBS launcher for Qwen HTTP server (no console) ---
$VbsLauncher = Join-Path $AppDir 'start_1c_bridge_silent.vbs'
$vbsContent = @"
CreateObject("Wscript.Shell").Run """$VenvPython"" ""$(Join-Path $AppDir 'mcp_server_1c_http.py')"" --port 8000", 0, False
"@
[System.IO.File]::WriteAllText($VbsLauncher, $vbsContent, [System.Text.ASCIIEncoding]::new())
Log "Created silent VBS launcher in $AppDir"

# List of supported MCP clients
# Each client: id, name, dir (under %APPDATA%), config filename, optional subdir
$MCPClients = @(
    @{ id = 'claude';   name = 'Claude Desktop';    dir = 'Claude';        config = 'claude_desktop_config.json' },
    @{ id = 'qwen';     name = 'Qwen Desktop';      dir = 'Qwen';          config = 'settings.json'; mcp_key = 'mcp_config' },
    @{ id = 'kimi';     name = 'Kimi Desktop';      dir = 'kimi-desktop';  config = 'mcp_config.json' },
    @{ id = 'reasonix'; name = 'Reasonix';          dir = 'reasonix';      config = '.mcp.json'; subdir = 'global-workspace' }
)

$ServerEntry = @{
    command = $VenvPython
    args    = @( (Join-Path $AppDir 'mcp_server_1c.py') )
    transportType = 'stdio'
    env     = @{
        ONEC_DATABASES_FILE = $DatabasesFile
    }
}

$ConfiguredClients = @()
$SkippedClients = @()

# Helper: recursively convert PSCustomObject to hashtable (PS 5.1 workaround)
function ConvertTo-HashtableDeep {
    param($obj)
    if ($null -eq $obj) { return $null }
    if ($obj -is [hashtable]) {
        $ht = @{}
        foreach ($k in $obj.Keys) { $ht[$k] = ConvertTo-HashtableDeep $obj[$k] }
        return $ht
    }
    if ($obj -is [Array]) {
        return @($obj | ForEach-Object { ConvertTo-HashtableDeep $_ })
    }
    if ($obj.GetType().Name -eq 'PSCustomObject') {
        $ht = @{}
        foreach ($p in $obj.PSObject.Properties) { $ht[$p.Name] = ConvertTo-HashtableDeep $p.Value }
        return $ht
    }
    return $obj
}

foreach ($client in $MCPClients) {
    try {
    Log "Processing $($client.name)..."

    # Allow path override via environment variable
    $envVarName = "ONEC_$($client.id.ToUpper())_CONFIG"
    $envPath = [Environment]::GetEnvironmentVariable($envVarName)
    if ($envPath) {
        $ConfigPath = $envPath
        $ConfigDir = Split-Path $ConfigPath -Parent
    } else {
        if ($client.subdir) {
            $ConfigDir  = Join-Path $UserAppData $client.dir | Join-Path -ChildPath $client.subdir
        } else {
            $ConfigDir  = Join-Path $UserAppData $client.dir
        }
        $ConfigPath = Join-Path $ConfigDir $client.config
    }

    # Detect if client is installed
    $clientInstalled = $false
    if (Test-Path $ConfigPath) {
        $clientInstalled = $true
        Log "$($client.name): found config $ConfigPath"
    } elseif (Test-Path $ConfigDir) {
        $clientInstalled = $true
        Log "$($client.name): found dir $ConfigDir (no config yet)"
    }

    if (-not $clientInstalled) {
        Log "$($client.name): not found - skipping."
        $SkippedClients += $client.name
        continue
    }

    # Create config dir if needed
    if (-not (Test-Path $ConfigDir)) {
        New-Item -ItemType Directory -Path $ConfigDir -Force | Out-Null
        Log "$($client.name): created dir $ConfigDir"
    }

    # Read or create config
    if (Test-Path $ConfigPath) {
        Log "$($client.name): updating existing config."
        try {
            $jsonText = Get-Content -Path $ConfigPath -Raw -Encoding UTF8
            $config   = $jsonText | ConvertFrom-Json
            $config   = ConvertTo-HashtableDeep $config
        } catch {
            Log "$($client.name): parse error - $($_.Exception.Message). Making backup."
            Copy-Item $ConfigPath ($ConfigPath + '.bak.' + (Get-Date -Format 'yyyyMMddHHmmss')) -Force
            $config = @{}
        }
    } else {
        Log "$($client.name): creating new config."
        $config = @{}
    }

    # Determine the JSON key for MCP servers (Qwen uses "mcp_config", others use "mcpServers")
    $mcpKey = if ($client.mcp_key) { $client.mcp_key } else { 'mcpServers' }

    if (-not $config.ContainsKey($mcpKey)) {
        $config[$mcpKey] = @{}
    }
    if ($config[$mcpKey] -isnot [hashtable]) {
        $tmp = @{}
        foreach ($p in $config[$mcpKey].PSObject.Properties) { $tmp[$p.Name] = $p.Value }
        $config[$mcpKey] = $tmp
    }

    # Qwen Desktop only allows npx/uvx — use npm launcher in ProgramData (no spaces!)
    if ($client.id -eq 'qwen') {
        $config[$mcpKey]['1c-bridge'] = @{
            command = 'npx'
            args    = @( $NpxDir )
            env     = @{
                ONEC_DATABASES_FILE = $DatabasesFile
            }
        }
    } else {
        $config[$mcpKey]['1c-bridge'] = $ServerEntry
    }

    $json = $config | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($ConfigPath, $json, [System.Text.UTF8Encoding]::new($false))
    Log "$($client.name): config written - $ConfigPath"
    $ConfiguredClients += $client.name

    } catch {
        Log "$($client.name): ERROR - $($_.Exception.Message)"
        $SkippedClients += "$($client.name) (error)"
    }
}

# Summary
if ($ConfiguredClients.Count -gt 0) {
    Log "Configured clients: $($ConfiguredClients -join ', ')"
} else {
    Log "WARNING: no MCP clients found."
    Log "1C Bridge is installed, but you need at least one MCP client:"
    foreach ($c in $MCPClients) {
        Log "  - $($c.name): download at $($c.id).ai or moonshot.cn"
    }
}

# -----------------------------------------------------------------------------
$bar = "=" * 70
Write-Host ""
Write-Host $bar -ForegroundColor Green
Write-Host "  УСТАНОВКА ЗАВЕРШЕНА УСПЕШНО" -ForegroundColor Green
Write-Host $bar -ForegroundColor Green
Write-Host ""
Write-Host "Окно закроется автоматически через 5 секунд." -ForegroundColor Yellow
Log "Установка завершена успешно."
Start-Sleep -Seconds 5
exit 0
