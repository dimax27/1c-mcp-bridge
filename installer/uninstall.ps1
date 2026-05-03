# =============================================================================
#  uninstall.ps1 — обратные действия при удалении.
#   • Убирает блок "1c-bridge" из claude_desktop_config.json
#     (остальные MCP-серверы пользователя не трогаем).
#   • COM-коннектор НЕ дерегистрируем — он мог быть нужен другим программам.
#   • Python и venv удалит сам Inno Setup вместе с папкой установки.
# =============================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'

# При запуске от админа $env:APPDATA указывает на профиль админа.
# Нам нужен профиль интерактивного пользователя — попробуем найти его.
function Get-InteractiveAppData {
    # Если запущен из uninstaller'а интерактивно — APPDATA уже верный
    if ($env:APPDATA -and (Test-Path (Join-Path $env:APPDATA 'Claude'))) {
        return $env:APPDATA
    }

    # Сканим C:\Users\*\AppData\Roaming\Claude
    $users = Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -notin @('Public','Default','Default User','All Users') }
    foreach ($u in $users) {
        $p = Join-Path $u.FullName 'AppData\Roaming'
        if (Test-Path (Join-Path $p 'Claude\claude_desktop_config.json')) {
            return $p
        }
    }
    return $null
}

$AppData = Get-InteractiveAppData
if (-not $AppData) {
    Write-Host "Конфиг Claude Desktop не найден — ничего удалять."
    exit 0
}

$ConfigPath = Join-Path $AppData 'Claude\claude_desktop_config.json'
if (-not (Test-Path $ConfigPath)) { exit 0 }

try {
    $jsonText = Get-Content -Path $ConfigPath -Raw -Encoding UTF8
    $config   = $jsonText | ConvertFrom-Json -AsHashtable

    if ($config.ContainsKey('mcpServers') -and $config.mcpServers -is [hashtable]) {
        if ($config.mcpServers.ContainsKey('1c-bridge')) {
            $config.mcpServers.Remove('1c-bridge')
            Write-Host "Удалён блок 1c-bridge из конфига Claude."

            $json = $config | ConvertTo-Json -Depth 10
            [System.IO.File]::WriteAllText($ConfigPath, $json, [System.Text.UTF8Encoding]::new($false))
        }
    }
} catch {
    Write-Host "Не удалось обновить конфиг: $($_.Exception.Message)"
    # Не падаем — деинсталлятор должен закончиться чисто.
}

exit 0
