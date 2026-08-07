<#
  build.ps1 — сборка установщика с запеканием версии в src/version.py.

  Единая точка сборки: локально (.\installer\build.ps1 -Version "1.2.3")
  и в GitHub Actions. Запекает версию в src/version.py (сервер показывает
  версию установщика), запускает Inno Setup и ВОССТАНАВЛИВАЕТ исходный
  version.py после сборки.

  Требует: Inno Setup 6 (обычно C:\Program Files (x86)\Inno Setup 6\ISCC.exe).
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$Version
)

$ErrorActionPreference = 'Stop'

# Версия обязана быть семвером (например 1.2.3 или 1.2.3-beta.1): она
# подставляется в src/version.py и в имя установщика.
if ($Version -notmatch '^[0-9]+\.[0-9]+\.[0-9]+(?:[-+][0-9A-Za-z.-]+)?$') {
    throw "Некорректная версия: '$Version' (ожидается x.y.z или x.y.z-метка)"
}

$installerDir = $PSScriptRoot
$root = Split-Path $PSScriptRoot -Parent
$versionFile = Join-Path $root 'src\version.py'

if (-not (Test-Path $versionFile)) {
    throw "Не найден $versionFile"
}

$original = [System.IO.File]::ReadAllText($versionFile)
$oldAppVersion = $env:APP_VERSION
# экранируем одинарные кавычки перед вставкой в Python-исходник
$safeVersion = $Version -replace "'", "''"

try {
    Write-Host "Запекаю версию $Version в src/version.py ..."
    [System.IO.File]::WriteAllText(
        $versionFile,
        "VERSION = '$safeVersion'`n",
        [System.Text.UTF8Encoding]::new($false)
    )

    # быстрая проверка, что сгенерированный файл компилируется
    & python -m py_compile $versionFile
    if ($LASTEXITCODE -ne 0) {
        throw "src/version.py не компилируется после запекания версии"
    }

    $env:APP_VERSION = $Version

    $iscc = "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
    if (-not (Test-Path $iscc)) {
        $iscc = "${env:ProgramFiles}\Inno Setup 6\ISCC.exe"
    }
    if (-not (Test-Path $iscc)) {
        throw "Не найден ISCC.exe (Inno Setup 6)"
    }

    & $iscc /Qp (Join-Path $installerDir 'setup.iss')
    if ($LASTEXITCODE -ne 0) {
        throw "Inno Setup завершился с кодом $LASTEXITCODE"
    }
}
finally {
    [System.IO.File]::WriteAllText(
        $versionFile,
        $original,
        [System.Text.UTF8Encoding]::new($false)
    )
    if ($null -eq $oldAppVersion) {
        Remove-Item Env:\APP_VERSION -ErrorAction SilentlyContinue
    } else {
        $env:APP_VERSION = $oldAppVersion
    }
}

$distDir = Join-Path $root 'dist'
$distPath = (Resolve-Path $distDir -ErrorAction Stop).Path
Write-Host "Готово: $distPath"
