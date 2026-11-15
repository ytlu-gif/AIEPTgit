#define MyAppName "AI互動式英語練習工具"
#define MyAppVersion "2.9.3"
#define MyAppExeName "AIEPT.exe"
#define ProjectDir SourcePath

[Setup]
AppId={{DB8C0430-CC06-4D58-B9E8-F13C7F0415C1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\{#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
PrivilegesRequiredOverridesAllowed=dialog
OutputDir={#ProjectDir}
OutputBaseFilename=AIEPT_Setup_v2.9.3
SetupIconFile={#ProjectDir}\app_icon.ico
SolidCompression=yes
WizardStyle=modern dynamic

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "chinesetraditional"; MessagesFile: "compiler:Languages\ChineseTraditional.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Required python-docx package template; keep user-generated DOCX files excluded below.
Source: "{#ProjectDir}\dist\AIEPT\_internal\docx\templates\default.docx"; DestDir: "{app}\_internal\docx\templates"; Flags: ignoreversion
; Public CA bundle needed by cloud speech; other PEM files remain excluded.
Source: "{#ProjectDir}\dist\AIEPT\_internal\certifi\cacert.pem"; DestDir: "{app}\_internal\certifi"; Flags: ignoreversion
Source: "{#ProjectDir}\dist\AIEPT\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#ProjectDir}\dist\AIEPT\*"; DestDir: "{app}"; Excludes: ".aiept-update-*.tmp,*.timings.json,*.timings.json.partial,config.json,chat_history.json,.env,.env.*,*.pem,*.key,*.p12,*.pfx,*.docx,*.mp3,*.wav,*.log"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,AI互動式英語練習工具}"; Flags: nowait postinstall skipifsilent
