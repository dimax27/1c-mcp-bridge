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

$installerDir = $PSScriptRoot
$root = Split-Path $PSScriptRoot -Parent
$versionFile = Join-Path $root 'src\version.py'

if (-not (Test-Path $versionFile)) {
    throw "Не найден $versionFile"
}

$original = [System.IO.File]::ReadAllText($versionFile)

try {
    Write-Host "Запекаю версию $Version в src/version.py ..."
    [System.IO.File]::WriteAllText(
        $versionFile,
        "VERSION = '$Version'`n",
        [System.Text.UTF8Encoding]::new($false)
    )

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
    Remove-Item Env:\APP_VERSION -ErrorAction SilentlyContinue
}

$distDir = Join-Path $root 'dist'
$distPath = (Resolve-Path $distDir -ErrorAction Stop).Path
Write-Host "Готово: $distPath"
