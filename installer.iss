; Inno Setup script for AudioSwitcher.
; Builds a per-user installer (no admin required) that registers the app
; in Windows' "Apps & Features" / installed-apps list, so it can be
; uninstalled the normal way instead of manually deleting files.
;
; Build with: "C:\Users\<you>\AppData\Local\Programs\Inno Setup 6\ISCC.exe" installer.iss
; (dist\AudioSwitcher.exe must already exist -- run PyInstaller first.)

#define MyAppName "AudioSwitcher"
#define MyAppVersion "1.7.0"
#define MyAppPublisher "Yuga Tanaka"
#define MyAppURL "https://github.com/CAIN218/audio-switcher"
#define MyAppExeName "AudioSwitcher.exe"

[Setup]
AppId={{D9F6B52A-3022-493B-B980-0EBA2C1EB140}
AppName={#MyAppName}
AppVerName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputBaseFilename=AudioSwitcher-Setup-v{#MyAppVersion}
OutputDir=installer_output
Compression=lzma2
SolidCompression=yes
SetupIconFile=app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\AudioSwitcher.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; postinstallはウィザードの完了画面のチェックボックスと連動する仕組みで、
; サイレントインストール時は(skipifsilentの有無に関わらず)発火しない。
; そのためサイレント時の再起動は[Code]のCurStepChanged側でWizardSilent()
; を見て明示的にExecしている(下記参照)。
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; アプリ自身が書き出す設定/ログもアンインストール時に一緒に消す
Type: files; Name: "{app}\app_config.txt"
Type: files; Name: "{app}\error.log"
; 「Start on boot」で作られるスタートアップフォルダのショートカット
; (winshell.startup()と{userstartup}は同じ場所を指す)もここで消す。
; アプリはインストーラー管理外のファイルなので、明示的に指定しないと
; アンインストール時に残ってしまう。
Type: files; Name: "{userstartup}\AudioSwitcher.lnk"

[Code]
var
  UpgradePage: TOutputMsgWizardPage;
  BackupConfigPath: String;
  RestoreStartupShortcut: Boolean;

procedure KillAudioSwitcher;
var
  ResultCode: Integer;
begin
  // 常駐したまま上書きインストール/アンインストールされると、実行中の
  // exeがロックされていて置き換え・削除に失敗する(PyInstallerのonefile
  // ビルドは親子2プロセスになるため、片方だけ終了しても残ることがある)。
  // インストール・アンインストールの両方で必ず終了させておく。
  Exec('taskkill.exe', '/F /IM AudioSwitcher.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

// AppIdは[Setup]のAppId={{D9F6B52A-3022-493B-B980-0EBA2C1EB140}と同じ値
// (Inno Setupが自動でこのGUID+"_is1"というキー名でアンインストール情報を
// HKCUに登録するため、ここでも同じ文字列を使って探しに行く)。
function GetUninstallString(): String;
var
  UninstallKey: String;
  Value: String;
begin
  UninstallKey := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{D9F6B52A-3022-493B-B980-0EBA2C1EB140}_is1';
  Value := '';
  RegQueryStringValue(HKCU, UninstallKey, 'UninstallString', Value);
  Result := Value;
end;

function IsUpgrade(): Boolean;
begin
  Result := (GetUninstallString() <> '');
end;

// 旧バージョンの実際のインストール先。ユーザーがディレクトリ選択ページで
// 前回と違うフォルダを選んだ場合、今回のセッションの{app}とは一致しない
// ことがあるため、app_config.txtの退避元は必ずレジストリのInstallLocation
// (旧バージョン自身が書き込んだ値)から求める。
function GetOldInstallLocation(): String;
var
  UninstallKey: String;
  Value: String;
begin
  UninstallKey := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{D9F6B52A-3022-493B-B980-0EBA2C1EB140}_is1';
  Value := '';
  RegQueryStringValue(HKCU, UninstallKey, 'InstallLocation', Value);
  Result := Value;
end;

procedure InitializeWizard();
begin
  // 既存バージョンが見つかった時だけ表示するお知らせページ
  // (見つからない場合はShouldSkipPageで自動的に飛ばす)。
  UpgradePage := CreateOutputMsgPage(wpWelcome,
    'Existing Installation Found',
    'A previous version of AudioSwitcher will be removed',
    'Setup has detected a previous version of AudioSwitcher already installed on this computer.'
    + #13#10 + #13#10 +
    'It will be automatically uninstalled before the new version is installed. Your device and hotkey settings will not be affected.');
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;
  if PageID = UpgradePage.ID then
    Result := not IsUpgrade();
end;

procedure RecreateStartupShortcut();
var
  WshShell: Variant;
  Shortcut: Variant;
begin
  // manage_startup()がPython側で作るショートカットと全く同じ内容
  // (Targetpath/Arguments/WorkingDirectory/IconLocation)で作り直す。
  try
    WshShell := CreateOleObject('WScript.Shell');
    Shortcut := WshShell.CreateShortcut(ExpandConstant('{userstartup}\AudioSwitcher.lnk'));
    Shortcut.TargetPath := ExpandConstant('{app}\{#MyAppExeName}');
    Shortcut.Arguments := '/background';
    Shortcut.WorkingDirectory := ExpandConstant('{app}');
    Shortcut.IconLocation := ExpandConstant('{app}\{#MyAppExeName}');
    Shortcut.Save;
  except
  end;
end;

procedure UninstallOldVersion();
var
  OldDir: String;
  UninstallString: String;
  ResultCode: Integer;
begin
  // 実際に使われている旧バージョンのアンインストーラーは[UninstallDelete]で
  // app_config.txtとスタートアップフォルダのショートカットも消してしまう。
  // アップグレードのつもりが設定やStart on bootが消えては困るので、実行前に
  // 退避しておき、新バージョンのインストール完了後(ssPostInstall)に戻す。
  BackupConfigPath := '';
  OldDir := GetOldInstallLocation();
  if (OldDir <> '') and FileExists(AddBackslash(OldDir) + 'app_config.txt') then
  begin
    BackupConfigPath := ExpandConstant('{tmp}\app_config.txt.bak');
    CopyFile(AddBackslash(OldDir) + 'app_config.txt', BackupConfigPath, False);
  end;
  RestoreStartupShortcut := FileExists(ExpandConstant('{userstartup}\AudioSwitcher.lnk'));

  UninstallString := GetUninstallString();
  if UninstallString <> '' then
  begin
    UninstallString := RemoveQuotes(UninstallString);
    Exec(UninstallString, '/VERYSILENT /SUPPRESSMSGBOX /NORESTART', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    // アンインストーラー(unins000.exe)は自分自身を一時フォルダにコピーして
    // 実際の削除処理をそちらに引き継ぐため、ewWaitUntilTerminatedで待てる
    // のは最初のプロセスの終了までで、スタートメニューのフォルダ削除等の
    // 実際のクリーンアップはその後も少し続く。これと後続の[Icons]作成が
    // 競合すると"IPersistFile::Save failed"で新バージョンのインストール
    // ごと失敗する(実際に発生した不具合)。少し待って確実に完了させる。
    Sleep(2000);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if (CurStep = ssInstall) and IsUpgrade() then
    UninstallOldVersion();

  if CurStep = ssPostInstall then
  begin
    if (BackupConfigPath <> '') and FileExists(BackupConfigPath) then
    begin
      CopyFile(BackupConfigPath, ExpandConstant('{app}\app_config.txt'), False);
      DeleteFile(BackupConfigPath);
    end;
    if RestoreStartupShortcut then
      RecreateStartupShortcut();

    // [Run]のpostinstallはサイレント時に発火しないため、Auto-updateから
    // /VERYSILENTで呼ばれた時だけここで明示的に起動し直す(ソフトウェア
    // アップデートのように無人で再起動する体験にするため)。
    if WizardSilent() then
      Exec(ExpandConstant('{app}\{#MyAppExeName}'), '', '', SW_SHOWNORMAL, ewNoWait, ResultCode);
  end;
end;

function InitializeSetup(): Boolean;
begin
  KillAudioSwitcher;
  Result := True;
end;

function InitializeUninstall(): Boolean;
begin
  KillAudioSwitcher;
  Result := True;
end;
