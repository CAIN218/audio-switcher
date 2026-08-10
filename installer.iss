; Inno Setup script for AudioSwitcher.
; Builds a per-user installer (no admin required) that registers the app
; in Windows' "Apps & Features" / installed-apps list, so it can be
; uninstalled the normal way instead of manually deleting files.
;
; Build with: "C:\Users\<you>\AppData\Local\Programs\Inno Setup 6\ISCC.exe" installer.iss
; (dist\AudioSwitcher.exe must already exist -- run PyInstaller first.)

#define MyAppName "AudioSwitcher"
#define MyAppVersion "1.2.2"
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
