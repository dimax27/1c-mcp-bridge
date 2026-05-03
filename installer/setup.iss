; =============================================================================
;  1C MCP Bridge — установщик
;  Связывает Claude Desktop с 1С:Предприятием через COM-коннектор и MCP.
;
;  Сборка:
;    iscc setup.iss
;  или через GitHub Actions при push'е тега v*.*.* — см. .github/workflows/build.yml
;
;  Требования компиляции: Inno Setup 6.2+ (Unicode).
; =============================================================================

#define MyAppName        "1C MCP Bridge"
#define MyAppNameSafe    "1cMcpBridge"
#define MyAppVersion     GetEnv("APP_VERSION")
#if MyAppVersion == ""
  #define MyAppVersion "0.1.0-dev"
#endif
#define MyAppPublisher   "Open Source"
#define MyAppURL         "https://github.com/dimax27/1c-mcp-bridge"
#define MyAppExeName     "mcp_server_1c.py"
#define PythonVersion    "3.12.7"
#define PythonInstaller  "python-" + PythonVersion + "-amd64.exe"
#define PythonDownloadURL "https://www.python.org/ftp/python/" + PythonVersion + "/" + PythonInstaller

[Setup]
AppId={{8B4C1A2E-9D3F-4E5A-B6C7-1F2D3E4A5B6C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
AppUpdatesURL={#MyAppURL}/releases
DefaultDirName={autopf}\{#MyAppNameSafe}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=..\LICENSE
OutputDir=..\dist
OutputBaseFilename=1cMcpBridge-Setup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
ArchitecturesAllowed=x64compatible
UninstallDisplayName={#MyAppName} {#MyAppVersion}
UninstallDisplayIcon={app}\assets\icon.ico
SetupLogging=yes

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Сам сервер и его зависимости
Source: "..\src\mcp_server_1c.py";       DestDir: "{app}";              Flags: ignoreversion
Source: "..\requirements.txt";           DestDir: "{app}";              Flags: ignoreversion
Source: "..\LICENSE";                    DestDir: "{app}";              Flags: ignoreversion
Source: "..\README.md";                  DestDir: "{app}";              Flags: ignoreversion

; PowerShell-скрипты установки (вызываются из [Run])
Source: "install.ps1";                   DestDir: "{app}\installer";    Flags: ignoreversion
Source: "test_connection.ps1";           DestDir: "{app}\installer";    Flags: ignoreversion
Source: "uninstall.ps1";                 DestDir: "{app}\installer";    Flags: ignoreversion
Source: "detect_1c.ps1";                 DestDir: "{app}\installer";    Flags: ignoreversion

; Иконка (опционально, см. assets/)
Source: "..\assets\icon.ico";            DestDir: "{app}\assets";       Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{group}\Открыть папку установки"; Filename: "{app}"
Name: "{group}\Удалить {#MyAppName}";    Filename: "{uninstallexe}"
Name: "{group}\Документация";            Filename: "{#MyAppURL}"

[Run]
; Финальная установка: Python (если нет), venv, зависимости, регистрация коннектора, конфиг Claude.
; Параметры передаются через переменные окружения, которые мы выставляем в [Code].
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{app}\installer\install.ps1"""; \
  Flags: runhidden waituntilterminated; \
  StatusMsg: "Устанавливаю Python и зависимости, регистрирую COM-коннектор, прописываю конфиг Claude Desktop..."

[UninstallRun]
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{app}\installer\uninstall.ps1"""; \
  Flags: runhidden waituntilterminated; \
  RunOnceId: "RemoveClaudeConfig"

; =============================================================================
;  Pascal Script: кастомные страницы мастера
; =============================================================================
[Code]
type
  TPlatformInfo = record
    Version:     String;       // "8.5.1.1150"
    InstallPath: String;       // "C:\Program Files\1cv8\8.5.1.1150"
    ProgID:      String;       // "V85.COMConnector"
    Bitness:     String;       // "x64" | "x86"
  end;

var
  // Найденные платформы 1С
  Platforms: array of TPlatformInfo;

  // Страница: выбор платформы
  PagePlatform:        TInputOptionWizardPage;

  // Страница: тип подключения
  PageConnectionMode:  TInputOptionWizardPage;

  // Страница: параметры файловой/серверной базы и аутентификации
  PageConnectionParams: TWizardPage;
  EditFileBasePath:     TNewEdit;
  EditServerName:       TNewEdit;
  EditRefName:          TNewEdit;
  EditUser:             TNewEdit;
  EditPassword:         TNewPasswordEdit;
  CheckOSAuth:          TNewCheckBox;
  LblFileBase:          TNewStaticText;
  LblServer, LblRef:    TNewStaticText;
  LblUser, LblPwd:      TNewStaticText;
  LblOSAuthHint:        TNewStaticText;

  // Страница: проверка соединения
  PageTest:             TWizardPage;
  BtnTest:              TNewButton;
  MemoTestOutput:       TNewMemo;
  TestPassed:           Boolean;

// ----------------------------------------------------------------------------
//  Сканирование реестра на установленные платформы 1С
// ----------------------------------------------------------------------------
procedure DetectPlatforms;
var
  RootKey:     Integer;
  ParentPath:  String;
  Subkeys:     TArrayOfString;
  i, n:        Integer;
  Ver, Path:   String;
  Major:       Integer;
  Info:        TPlatformInfo;
begin
  SetArrayLength(Platforms, 0);

  // 1С пишется в HKLM\SOFTWARE\1C\1Cv8\<version>\... — для x64-системы оба варианта,
  // включая Wow6432Node, на всякий случай.
  RootKey := HKEY_LOCAL_MACHINE;

  if RegGetSubkeyNames(RootKey, 'SOFTWARE\1C\1Cv8', Subkeys) then
  begin
    n := GetArrayLength(Subkeys);
    for i := 0 to n - 1 do
    begin
      Ver  := Subkeys[i];
      Path := '';
      RegQueryStringValue(RootKey, 'SOFTWARE\1C\1Cv8\' + Ver, 'InstalledLocation', Path);
      if Path = '' then
        RegQueryStringValue(RootKey, 'SOFTWARE\1C\1Cv8\' + Ver, 'Path', Path);

      // Major-версия: "8.3.24.1234" -> 83, "8.5.1.1150" -> 85
      if (Length(Ver) >= 3) and (Ver[1] = '8') then
      begin
        Major := StrToIntDef(Ver[1] + Ver[3], 0);
        if Major >= 82 then
        begin
          Info.Version     := Ver;
          Info.InstallPath := Path;
          Info.ProgID      := 'V' + IntToStr(Major) + '.COMConnector';
          Info.Bitness     := 'x64'; // эвристика — рассмотрим только 64-битные ветки
          SetArrayLength(Platforms, GetArrayLength(Platforms) + 1);
          Platforms[GetArrayLength(Platforms) - 1] := Info;
        end;
      end;
    end;
  end;
end;

// ----------------------------------------------------------------------------
//  Создание кастомных страниц
// ----------------------------------------------------------------------------
procedure CreatePlatformPage;
var
  i: Integer;
  Caption, Description: String;
begin
  if GetArrayLength(Platforms) = 0 then
  begin
    PagePlatform := CreateInputOptionPage(wpLicense,
      'Платформа 1С:Предприятие',
      'Платформа 1С на компьютере не обнаружена',
      'Установщик не нашёл зарегистрированных версий 1С:Предприятия. ' +
      'Без установленной платформы COM-коннектор не сможет подключиться к информационной базе.' + #13#10 + #13#10 +
      'Вы можете продолжить установку, но перед использованием потребуется установить платформу 1С 8.2 или новее.',
      False, False);
    PagePlatform.Add('Продолжить без установленной 1С');
    PagePlatform.SelectedValueIndex := 0;
    Exit;
  end;

  PagePlatform := CreateInputOptionPage(wpLicense,
    'Платформа 1С:Предприятие',
    'Выберите версию платформы для подключения',
    'Установщик нашёл следующие версии 1С на этом компьютере. ' +
    'Будет зарегистрирован COM-коннектор именно для выбранной версии.',
    True, False);

  for i := 0 to GetArrayLength(Platforms) - 1 do
  begin
    Caption := '1С:Предприятие ' + Platforms[i].Version + '   (' + Platforms[i].ProgID + ')';
    PagePlatform.Add(Caption);
  end;
  // По умолчанию выбираем самую старшую версию (последняя в списке реестра обычно).
  PagePlatform.SelectedValueIndex := GetArrayLength(Platforms) - 1;
end;

procedure CreateConnectionModePage;
begin
  PageConnectionMode := CreateInputOptionPage(PagePlatform.ID,
    'Тип информационной базы',
    'Файловая или клиент-серверная база',
    'Укажите, как развёрнута ваша информационная база 1С.',
    True, False);
  PageConnectionMode.Add('Файловая база (.1CD на диске)');
  PageConnectionMode.Add('Клиент-серверная база (1С-сервер + СУБД)');
  PageConnectionMode.SelectedValueIndex := 1;
end;

procedure UpdateAuthFields(Sender: TObject); forward;

procedure CreateConnectionParamsPage;
var
  Y: Integer;
begin
  PageConnectionParams := CreateCustomPage(PageConnectionMode.ID,
    'Параметры подключения',
    'Введите данные для подключения к информационной базе 1С.');

  Y := 0;

  // --- Файловая база ---
  LblFileBase := TNewStaticText.Create(PageConnectionParams);
  LblFileBase.Parent := PageConnectionParams.Surface;
  LblFileBase.Caption := 'Путь к каталогу файловой базы (.1CD):';
  LblFileBase.Top := Y;
  LblFileBase.Width := PageConnectionParams.SurfaceWidth;
  Y := Y + LblFileBase.Height + 4;

  EditFileBasePath := TNewEdit.Create(PageConnectionParams);
  EditFileBasePath.Parent := PageConnectionParams.Surface;
  EditFileBasePath.Top := Y;
  EditFileBasePath.Width := PageConnectionParams.SurfaceWidth;
  EditFileBasePath.Text := 'C:\1C\bases\MyBase';
  Y := Y + EditFileBasePath.Height + 16;

  // --- Сервер ---
  LblServer := TNewStaticText.Create(PageConnectionParams);
  LblServer.Parent := PageConnectionParams.Surface;
  LblServer.Caption := 'Адрес сервера 1С (Srvr), например 127.0.0.1 или srv-1c:1541:';
  LblServer.Top := Y;
  LblServer.Width := PageConnectionParams.SurfaceWidth;
  Y := Y + LblServer.Height + 4;

  EditServerName := TNewEdit.Create(PageConnectionParams);
  EditServerName.Parent := PageConnectionParams.Surface;
  EditServerName.Top := Y;
  EditServerName.Width := PageConnectionParams.SurfaceWidth;
  EditServerName.Text := '127.0.0.1';
  Y := Y + EditServerName.Height + 12;

  // --- Имя ИБ ---
  LblRef := TNewStaticText.Create(PageConnectionParams);
  LblRef.Parent := PageConnectionParams.Surface;
  LblRef.Caption := 'Имя информационной базы в кластере (Ref):';
  LblRef.Top := Y;
  LblRef.Width := PageConnectionParams.SurfaceWidth;
  Y := Y + LblRef.Height + 4;

  EditRefName := TNewEdit.Create(PageConnectionParams);
  EditRefName.Parent := PageConnectionParams.Surface;
  EditRefName.Top := Y;
  EditRefName.Width := PageConnectionParams.SurfaceWidth;
  Y := Y + EditRefName.Height + 16;

  // --- OS-аутентификация ---
  CheckOSAuth := TNewCheckBox.Create(PageConnectionParams);
  CheckOSAuth.Parent := PageConnectionParams.Surface;
  CheckOSAuth.Caption := 'Аутентификация средствами Windows (текущий пользователь)';
  CheckOSAuth.Top := Y;
  CheckOSAuth.Width := PageConnectionParams.SurfaceWidth;
  CheckOSAuth.Checked := True;
  CheckOSAuth.OnClick := @UpdateAuthFields;
  Y := Y + CheckOSAuth.Height + 4;

  LblOSAuthHint := TNewStaticText.Create(PageConnectionParams);
  LblOSAuthHint.Parent := PageConnectionParams.Surface;
  LblOSAuthHint.Caption := 'При снятой галочке введите логин и пароль пользователя 1С.';
  LblOSAuthHint.Top := Y;
  LblOSAuthHint.Width := PageConnectionParams.SurfaceWidth;
  Y := Y + LblOSAuthHint.Height + 12;

  // --- Логин/Пароль ---
  LblUser := TNewStaticText.Create(PageConnectionParams);
  LblUser.Parent := PageConnectionParams.Surface;
  LblUser.Caption := 'Логин пользователя 1С:';
  LblUser.Top := Y;
  LblUser.Width := PageConnectionParams.SurfaceWidth;
  Y := Y + LblUser.Height + 4;

  EditUser := TNewEdit.Create(PageConnectionParams);
  EditUser.Parent := PageConnectionParams.Surface;
  EditUser.Top := Y;
  EditUser.Width := PageConnectionParams.SurfaceWidth;
  Y := Y + EditUser.Height + 8;

  LblPwd := TNewStaticText.Create(PageConnectionParams);
  LblPwd.Parent := PageConnectionParams.Surface;
  LblPwd.Caption := 'Пароль:';
  LblPwd.Top := Y;
  LblPwd.Width := PageConnectionParams.SurfaceWidth;
  Y := Y + LblPwd.Height + 4;

  EditPassword := TNewPasswordEdit.Create(PageConnectionParams);
  EditPassword.Parent := PageConnectionParams.Surface;
  EditPassword.Top := Y;
  EditPassword.Width := PageConnectionParams.SurfaceWidth;
end;

procedure UpdateAuthFields(Sender: TObject);
var
  E: Boolean;
begin
  E := not CheckOSAuth.Checked;
  EditUser.Enabled := E;
  EditPassword.Enabled := E;
  LblUser.Enabled := E;
  LblPwd.Enabled := E;
end;

procedure UpdateConnectionFieldsVisibility;
var
  IsServer: Boolean;
begin
  IsServer := (PageConnectionMode.SelectedValueIndex = 1);
  LblFileBase.Visible      := not IsServer;
  EditFileBasePath.Visible := not IsServer;
  LblServer.Visible := IsServer;  EditServerName.Visible := IsServer;
  LblRef.Visible    := IsServer;  EditRefName.Visible    := IsServer;
end;

procedure CreateTestPage;
begin
  PageTest := CreateCustomPage(PageConnectionParams.ID,
    'Проверка подключения',
    'Можно сразу убедиться, что 1С отзывается на введённые параметры.');

  BtnTest := TNewButton.Create(PageTest);
  BtnTest.Parent := PageTest.Surface;
  BtnTest.Caption := 'Проверить подключение к 1С';
  BtnTest.Top := 0;
  BtnTest.Width := 240;
  BtnTest.Height := 28;
  BtnTest.OnClick := @TestConnectionClick;

  MemoTestOutput := TNewMemo.Create(PageTest);
  MemoTestOutput.Parent := PageTest.Surface;
  MemoTestOutput.Top := BtnTest.Top + BtnTest.Height + 12;
  MemoTestOutput.Width := PageTest.SurfaceWidth;
  MemoTestOutput.Height := PageTest.SurfaceHeight - BtnTest.Height - 12;
  MemoTestOutput.ScrollBars := ssVertical;
  MemoTestOutput.ReadOnly := True;
  MemoTestOutput.Text := 'Нажмите «Проверить» — установщик попытается подключиться к 1С с введёнными параметрами.' + #13#10 +
                        '(Этот шаг не обязателен, можно пропустить кнопкой «Далее».)';
end;

// ----------------------------------------------------------------------------
//  Сборка строки соединения и параметров для install.ps1
// ----------------------------------------------------------------------------
function BuildConnectionString: String;
var
  Auth: String;
begin
  if CheckOSAuth.Checked then
    Auth := ''
  else
    Auth := ';Usr="' + EditUser.Text + '";Pwd="' + EditPassword.Text + '"';

  if PageConnectionMode.SelectedValueIndex = 0 then
    Result := 'File="' + EditFileBasePath.Text + '"' + Auth
  else
    Result := 'Srvr="' + EditServerName.Text + '";Ref="' + EditRefName.Text + '"' + Auth;
end;

function GetSelectedProgID: String;
begin
  if GetArrayLength(Platforms) > 0 then
    Result := Platforms[PagePlatform.SelectedValueIndex].ProgID
  else
    Result := 'V85.COMConnector';
end;

function GetSelectedDllPath: String;
begin
  if GetArrayLength(Platforms) > 0 then
    Result := Platforms[PagePlatform.SelectedValueIndex].InstallPath + '\bin\comcntr.dll'
  else
    Result := '';
end;

// ----------------------------------------------------------------------------
//  Тест подключения (вызов test_connection.ps1)
// ----------------------------------------------------------------------------
procedure TestConnectionClick(Sender: TObject);
var
  ResultCode: Integer;
  TmpFile, Cmd, Output: String;
begin
  MemoTestOutput.Lines.Clear;
  MemoTestOutput.Lines.Add('Запускаю проверку...');
  MemoTestOutput.Lines.Add('');

  TmpFile := ExpandConstant('{tmp}\test_output.txt');

  // Кладём parameters во временный файл — чтобы не возиться с экранированием на cmd-line
  SaveStringToFile(ExpandConstant('{tmp}\test_params.txt'),
    'PROGID=' + GetSelectedProgID + #13#10 +
    'CONNSTR=' + BuildConnectionString + #13#10 +
    'DLLPATH=' + GetSelectedDllPath + #13#10, False);

  Cmd := '-ExecutionPolicy Bypass -NoProfile -File "' +
         ExpandConstant('{src}\test_connection.ps1') + '" -ParamsFile "' +
         ExpandConstant('{tmp}\test_params.txt') + '" -OutputFile "' + TmpFile + '"';

  if Exec('powershell.exe', Cmd, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
  begin
    LoadStringFromFile(TmpFile, Output);
    MemoTestOutput.Lines.Text := Output;
    TestPassed := (ResultCode = 0);
  end
  else
  begin
    MemoTestOutput.Lines.Add('Не удалось запустить PowerShell.');
    TestPassed := False;
  end;
end;

// ----------------------------------------------------------------------------
//  Точки входа Inno Setup
// ----------------------------------------------------------------------------
procedure InitializeWizard;
begin
  DetectPlatforms;
  CreatePlatformPage;
  CreateConnectionModePage;
  CreateConnectionParamsPage;
  CreateTestPage;
  TestPassed := False;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = PageConnectionParams.ID then
    UpdateConnectionFieldsVisibility;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;

  if CurPageID = PageConnectionParams.ID then
  begin
    if PageConnectionMode.SelectedValueIndex = 0 then
    begin
      if Trim(EditFileBasePath.Text) = '' then
      begin
        MsgBox('Укажите путь к файловой базе.', mbError, MB_OK);
        Result := False;
      end;
    end
    else
    begin
      if (Trim(EditServerName.Text) = '') or (Trim(EditRefName.Text) = '') then
      begin
        MsgBox('Заполните адрес сервера и имя информационной базы.', mbError, MB_OK);
        Result := False;
      end;
    end;
    if Result and (not CheckOSAuth.Checked) and (Trim(EditUser.Text) = '') then
    begin
      MsgBox('Укажите логин 1С или включите Windows-аутентификацию.', mbError, MB_OK);
      Result := False;
    end;
  end;
end;

// Передаём параметры в install.ps1 через переменные окружения [Run].
// Inno Setup expandconstant'ит {code:...} прямо в Parameters/Filename, но проще
// сохранить файл и читать из install.ps1.
procedure CurStepChanged(CurStep: TSetupStep);
var
  ParamsPath: String;
begin
  if CurStep = ssInstall then
  begin
    ParamsPath := ExpandConstant('{app}\installer\install_params.txt');
    ForceDirectories(ExpandConstant('{app}\installer'));
    SaveStringToFile(ParamsPath,
      'PROGID='   + GetSelectedProgID    + #13#10 +
      'CONNSTR='  + BuildConnectionString + #13#10 +
      'DLLPATH='  + GetSelectedDllPath   + #13#10 +
      'APPDIR='   + ExpandConstant('{app}') + #13#10 +
      'USERAPPDATA=' + ExpandConstant('{userappdata}') + #13#10,
      False);
  end;
end;
