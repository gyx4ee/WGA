#define MyAppName "WinSys Guardian Advanced"
#define MyAppVersion "0.1.1"
#define MyAppPublisher "WGA"
#define MyAppExeName "WGA.exe"
#define MySourceRoot "C:\Users\PC\Documents\New project"
#define MyDistRoot "C:\Users\PC\Documents\New project\dist\WGA"

[Setup]
AppId={{9C7B8F39-6F12-40BC-8D7F-0AE92C291F72}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\WinSys Guardian Advanced
DefaultGroupName=WinSys Guardian Advanced
DisableProgramGroupPage=yes
DisableDirPage=no
LicenseFile=
OutputDir={#MySourceRoot}\installer-output
OutputBaseFilename=WGA-Setup
SetupIconFile={#MySourceRoot}\assets\wga-icon.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Types]
Name: "compact"; Description: "Application only - installers can be downloaded later"
Name: "custom"; Description: "Custom installation"; Flags: iscustom

[Components]
Name: "main"; Description: "WinSys Guardian Advanced application"; Types: compact custom; Flags: fixed

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "{#MyDistRoot}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: main
Source: "{#MySourceRoot}\installers_manifest.json"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "{#MySourceRoot}\version.json"; DestDir: "{app}"; Flags: ignoreversion; Components: main

[Icons]
Name: "{group}\WinSys Guardian Advanced"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall WinSys Guardian Advanced"; Filename: "{uninstallexe}"
Name: "{autodesktop}\WinSys Guardian Advanced"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch WinSys Guardian Advanced"; Flags: nowait postinstall skipifsilent

[Code]
procedure InitializeWizard;
begin
  WizardForm.SelectDirLabel.Caption :=
    'Choose where to install WinSys Guardian Advanced. You can select a local SSD/HDD folder, a flash drive, or another writable path.';
end;
