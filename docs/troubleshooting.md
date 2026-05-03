# Диагностика проблем

## Установка

### «Class not registered» при тесте подключения

COM-коннектор не зарегистрирован для выбранной версии 1С. Проверь:

```powershell
[Type]::GetTypeFromProgID("V85.COMConnector")  # или V83/V82 — твоя версия
```

Если `$null` — выполни от админа:

```powershell
regsvr32 "C:\Program Files\1cv8\<твоя_версия>\bin\comcntr.dll"
```

### `pywin32` падает при `pip install` («Building wheel error»)

Под текущей версией Python нет готового wheel'а. Решение — установить
Python 3.12 (под него wheel есть всегда).

### Установщик не нашёл платформу 1С

Платформа установлена, но реестр пуст. Так бывает при ручном копировании
без официального инсталлятора. Запусти диагностику:

```powershell
powershell -File "C:\Program Files\1cMcpBridge\installer\detect_1c.ps1"
```

Если 1С нашлась, но `COMRegistered=нет` — выполни `regsvr32` вручную (см. выше).

### «Не удалось подключиться к серверу 1С»

* Имя сервера должно совпадать с тем, что у тебя в окне выбора ИБ клиента 1С.
  Часто это `имя_компьютера`, иногда с портом: `srv-1c:1541`.
* Имя ИБ — это её имя в кластере, не имя конфигурации внутри. Например, в
  кластере база называется `БП3`, а конфигурация в ней — `Бухгалтерия предприятия`.
* Проверь, есть ли свободные клиентские лицензии — COM-сессия их использует.

## Работа

### MCP-сервер не появляется в Claude Desktop

1. Полностью закрой Claude Desktop через трей: правый клик → Quit. Просто
   закрытие окна не выгружает приложение из памяти.
2. Запусти заново.
3. Если не помогло — проверь лог:

   ```powershell
   Get-Content "$env:APPDATA\Claude\logs\mcp-server-1c-bridge.log"
   ```

   Чаще всего видна точная ошибка Python.

### Конфиг сломался / нужно начать заново

```powershell
# Backup'нуть и удалить блок
Copy-Item "$env:APPDATA\Claude\claude_desktop_config.json" "$env:APPDATA\Claude\claude_desktop_config.json.bak"
notepad "$env:APPDATA\Claude\claude_desktop_config.json"
```

В файле удали ключ `"1c-bridge"` из `mcpServers`. Потом перезапусти инсталлятор —
он впишет блок заново с теми параметрами, что укажешь в мастере.

### Запросы выполняются медленно или падают по таймауту

* По умолчанию лимит 1000 строк в результате (можно увеличить вызовом
  `execute_query` с `limit`).
* Если запрос трогает большие регистры без отбора по периоду — добавь параметры
  `&НачДата` и `&КонДата` в виртуальную таблицу.
* Большой результат лучше агрегировать в самом запросе (СГРУППИРОВАТЬ ПО,
  СУММА, ВЫБРАТЬ ПЕРВЫЕ N), а не тащить сырьё.

## Разработка

### Локальный smoke-test без инсталлятора

```powershell
cd path\to\1c-mcp-bridge
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:ONEC_COMCONNECTOR_PROGID = "V85.COMConnector"
$env:ONEC_CONNECTION_STRING   = 'Srvr="127.0.0.1";Ref="МояБаза"'
python -c "from src.mcp_server_1c import get_connection; print(get_connection().Метаданные.Имя)"
```

### Сборка инсталлятора локально

```powershell
$env:APP_VERSION = '0.1.0'
& "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe" installer\setup.iss
# .\dist\1cMcpBridge-Setup-0.1.0.exe
```
