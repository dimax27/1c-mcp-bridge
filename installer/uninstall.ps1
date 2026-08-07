# =============================================================================
#  uninstall.ps1 — cleanup during uninstall.
#   * Removes "1c-bridge" block from ALL supported MCP client configs:
#     ChatGPT Desktop (Codex), Claude Desktop, Qwen Desktop, Kimi Desktop,
#     Reasonix.
#   * ChatGPT (Codex config.toml) обрабатывается ВСЕГДА и независимо от
#     остальных клиентов: даже если других клиентов нет, мёртвая секция
#     из ~/.codex/config.toml будет удалена.
#   * COM connector is NOT unregistered (other apps may need it).
#   * Python and venv are removed by Inno Setup along with the install folder.
# =============================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'

# Корень установки: uninstall.ps1 лежит в {app}\installer
$InstallRoot = Split-Path $PSScriptRoot -Parent
$VenvPython = Join-Path $InstallRoot '.venv\Scripts\python.exe'
$ClientsScript = Join-Path $InstallRoot 'clients_config.py'

# Останавливаем HTTP-сервер моста, если он ещё запущен (иначе он держит
# файлы venv и порт 8000 на время удаления).
$StopServerScript = Join-Path $PSScriptRoot 'stop_http_server.ps1'
if (Test-Path $StopServerScript) {
    try {
        . $StopServerScript
        Stop-BridgeHttpServer | Out-Null
    } catch {
        Write-Host "Не удалось остановить HTTP-сервер моста: $($_.Exception.Message)"
    }
}

# --- ChatGPT Desktop (OpenAI Codex): config.toml лежит в профиле пользователя.
#     Обрабатываем до любых ранних выходов: на машине только с ChatGPT
#     деинсталлятор обязан удалить секцию 1c-bridge. ---
$UserProfiles = Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -notin @('Public','Default','Default User','All Users') }
foreach ($u in $UserProfiles) {
    $CodexCfg = Join-Path $u.FullName '.codex\config.toml'
    if (-not (Test-Path $CodexCfg)) { continue }
    if ((Test-Path $VenvPython) -and (Test-Path $ClientsScript)) {
        try {
            & $VenvPython $ClientsScript remove-codex --path $CodexCfg --server '1c-bridge' | Out-Null
            Write-Host "Removed 1c-bridge from ChatGPT Desktop config: $CodexCfg"
        } catch {
            Write-Host "Failed to update ChatGPT Desktop config: $($_.Exception.Message)"
        }
    } else {
        Write-Host "WARNING: секция 1c-bridge осталась в $CodexCfg — удалите её вручную (мост уже удалён)."
    }
}

# When running as admin, $env:APPDATA points to admin's profile.
# We need the interactive user's profile — try to find it.
function Get-InteractiveAppData {
    if ($env:APPDATA) {
        $clients = @('Claude', 'Qwen', 'kimi-desktop', 'reasonix')
        foreach ($c in $clients) {
            if (Test-Path (Join-Path $env:APPDATA $c)) { return $env:APPDATA }
        }
    }

    # Scan C:\Users\*\AppData\Roaming for any matching client configs
    $users = Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -notin @('Public','Default','Default User','All Users') }
    foreach ($u in $users) {
        $p = Join-Path $u.FullName 'AppData\Roaming'
        foreach ($c in @('Claude\claude_desktop_config.json', 'Qwen\settings.json', 'kimi-desktop\mcp_config.json', 'reasonix\global-workspace\.mcp.json')) {
            if (Test-Path (Join-Path $p $c)) { return $p }
        }
    }
    return $null
}

# Supported MCP clients: dir, config filename
$MCPClients = @(
    @{ dir = 'Claude';        config = 'claude_desktop_config.json' },
    @{ dir = 'Qwen';          config = 'settings.json'; mcp_key = 'mcp_config' },
    @{ dir = 'kimi-desktop';  config = 'mcp_config.json' },
    @{ dir = 'reasonix';      config = '.mcp.json'; subdir = 'global-workspace' }
)


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

$AppData = Get-InteractiveAppData
if (-not $AppData) {
    Write-Host "No other MCP client configs found — nothing to remove."
    exit 0
}

foreach ($client in $MCPClients) {
    if ($client.subdir) {
        $ConfigPath = Join-Path $AppData ($client.dir + '\' + $client.subdir + '\' + $client.config)
    } else {
        $ConfigPath = Join-Path $AppData ($client.dir + '\' + $client.config)
    }
    if (-not (Test-Path $ConfigPath)) { continue }

    try {
        $jsonText = Get-Content -Path $ConfigPath -Raw -Encoding UTF8
        $config   = $jsonText | ConvertFrom-Json
        $config   = ConvertTo-HashtableDeep $config

        $mcpKey = if ($client.mcp_key) { $client.mcp_key } else { 'mcpServers' }

        if ($config.ContainsKey($mcpKey) -and $config[$mcpKey] -is [hashtable]) {
            if ($config[$mcpKey].ContainsKey('1c-bridge')) {
                $config[$mcpKey].Remove('1c-bridge')
                Write-Host "Removed 1c-bridge from $($client.dir) config."

                $json = $config | ConvertTo-Json -Depth 10
                [System.IO.File]::WriteAllText($ConfigPath, $json, [System.Text.UTF8Encoding]::new($false))
            }
        }
    } catch {
        Write-Host "Failed to update $($client.dir) config: $($_.Exception.Message)"
    }
}

exit 0
