
#define MyAppName "Satış Rapor Güncelleme"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Şirketiniz"
#define MyAppExeName "SatisRaporGuncelleme.exe"
[Setup]
AppId={{6C8F7E2E-A4B2-49F3-8E8F-1A6A55C4A0B7}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputBaseFilename=SatisRaporGuncelleme-Installer
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
; SetupIconFile=icon.ico
[Languages]
Name: "tr"; MessagesFile: "compiler:Languages\Turkish.isl"
[Files]
Source: "dist\SatisRaporGuncelleme\*"; DestDir: "{app}"; Flags: recursesubdirs
[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram, {#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
