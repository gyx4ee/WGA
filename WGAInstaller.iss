#define MyAppName "WinSys Guardian Advanced"
#define MyAppVersion "0.2.42"
#define MyAppPublisher "WGA"
#define MyAppExeName "WGA.exe"
#define MySourceRoot "C:\Users\PC\Documents\New project"
#define MyDistRoot "C:\Users\PC\Documents\New project\dist\WGA"

[Setup]
AppId={{9C7B8F39-6F12-40BC-8D7F-0AE92C291F72}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={code:GetSuggestedInstallDir}
DefaultGroupName=WinSys Guardian Advanced
DisableProgramGroupPage=yes
DisableDirPage=no
UsePreviousAppDir=no
OutputDir={#MySourceRoot}\installer-output
OutputBaseFilename=WGA-Setup-USB-0.2.42
SetupIconFile={#MySourceRoot}\assets\wga-icon.ico
Compression=lzma2/normal
SolidCompression=no
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Types]
Name: "usbonly"; Description: "Install only on USB external storage"; Flags: iscustom

[Components]
Name: "main"; Description: "WinSys Guardian Advanced application"; Types: usbonly; Flags: fixed

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "{#MyDistRoot}\*"; DestDir: "{app}"; Excludes: "Installers\*"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: main
Source: "{#MySourceRoot}\installers_manifest.json"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "{#MySourceRoot}\version.json"; DestDir: "{app}"; Flags: ignoreversion; Components: main

[Icons]
Name: "{group}\WinSys Guardian Advanced"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall WinSys Guardian Advanced"; Filename: "{uninstallexe}"
Name: "{autodesktop}\WinSys Guardian Advanced"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch WinSys Guardian Advanced"; Flags: nowait postinstall skipifsilent

[Code]
const
  DRIVE_REMOVABLE = 2;
  DRIVE_FIXED = 3;

var
  AllowedDriveRoots: string;

function GetDriveType(lpRootPathName: string): UINT;
  external 'GetDriveTypeW@kernel32.dll stdcall';

function NormalizeDriveRoot(const DirName: string): string;
begin
  Result := AddBackslash(Uppercase(ExtractFileDrive(DirName)));
end;

procedure AddAllowedDriveRoot(const DriveRoot: string);
var
  NormalizedRoot: string;
begin
  NormalizedRoot := NormalizeDriveRoot(DriveRoot);
  if NormalizedRoot = '' then
    Exit;

  if Pos('|' + NormalizedRoot + '|', AllowedDriveRoots) = 0 then
    AllowedDriveRoots := AllowedDriveRoots + '|' + NormalizedRoot + '|';
end;

procedure AddRemovableUsbDrives;
var
  DriveCode: Integer;
  DriveRoot: string;
begin
  for DriveCode := Ord('D') to Ord('Z') do
  begin
    DriveRoot := Chr(DriveCode) + ':\';
    if DirExists(DriveRoot) and (GetDriveType(DriveRoot) = DRIVE_REMOVABLE) then
      AddAllowedDriveRoot(DriveRoot);
  end;
end;

procedure AddUsbFixedDrives;
var
  TempFileName: string;
  CommandLine: string;
  ResultCode: Integer;
  DriveListText: AnsiString;
  LineStart: Integer;
  LineEnd: Integer;
  CurrentLine: string;
begin
  TempFileName := ExpandConstant('{tmp}\wga_usb_drives.txt');
  DeleteFile(TempFileName);

  CommandLine :=
    '/C powershell -NoProfile -ExecutionPolicy Bypass -Command ' +
    '"$letters = @(Get-Disk | Where-Object { $_.BusType.ToString() -eq ''USB'' } | ' +
    'ForEach-Object { Get-Partition -DiskNumber $_.Number -ErrorAction SilentlyContinue } | ' +
    'Where-Object { $_.DriveLetter } | ForEach-Object { ($_.DriveLetter + '':'' ) }); ' +
    '$letters | Sort-Object -Unique | Set-Content -Encoding ASCII ''' + TempFileName + '''"';

  if not Exec(ExpandConstant('{cmd}'), CommandLine, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    Exit;

  if (ResultCode <> 0) or (not LoadStringFromFile(TempFileName, DriveListText)) then
    Exit;

  LineStart := 1;
  while LineStart <= Length(DriveListText) do
  begin
    LineEnd := LineStart;
    while (LineEnd <= Length(DriveListText)) and
      (DriveListText[LineEnd] <> #13) and (DriveListText[LineEnd] <> #10) do
      LineEnd := LineEnd + 1;

    CurrentLine := Trim(Copy(DriveListText, LineStart, LineEnd - LineStart));
    if CurrentLine <> '' then
      AddAllowedDriveRoot(CurrentLine);

    while (LineEnd <= Length(DriveListText)) and
      ((DriveListText[LineEnd] = #13) or (DriveListText[LineEnd] = #10)) do
      LineEnd := LineEnd + 1;

    LineStart := LineEnd;
  end;
end;

procedure LoadAllowedDriveRoots;
begin
  AllowedDriveRoots := '';
  AddRemovableUsbDrives;
  AddUsbFixedDrives;
end;

function HasAllowedUsbDrive: Boolean;
begin
  Result := AllowedDriveRoots <> '';
end;

function IsAllowedInstallDrive(const DriveRoot: string): Boolean;
begin
  Result := Pos('|' + NormalizeDriveRoot(DriveRoot) + '|', AllowedDriveRoots) > 0;
end;

function FindFirstAllowedInstallDir: string;
var
  DriveCode: Integer;
  DriveRoot: string;
begin
  Result := '';
  for DriveCode := Ord('D') to Ord('Z') do
  begin
    DriveRoot := Chr(DriveCode) + ':\';
    if IsAllowedInstallDrive(DriveRoot) then
    begin
      Result := DriveRoot + 'WinSys Guardian Advanced';
      Exit;
    end;
  end;
end;

function GetSuggestedInstallDir(Param: string): string;
begin
  Result := FindFirstAllowedInstallDir();
  if Result = '' then
    Result := 'E:\WinSys Guardian Advanced';
end;

function InitializeSetup(): Boolean;
begin
  LoadAllowedDriveRoots();
  if HasAllowedUsbDrive() then
  begin
    Result := True;
    Exit;
  end;

  MsgBox(
    'Ne e otkrito podhodiashto USB ustroystvo.' + #13#10#13#10 +
    'Razreshena e samo instalaciya vurhu USB flashka, USB HDD ili USB SSD.' + #13#10 +
    'Lokalna instalaciya i SATA diskove ne sa pozvoleni.',
    mbError,
    MB_OK
  );
  Result := False;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  SelectedRoot: string;
begin
  Result := True;

  if CurPageID = wpSelectDir then
  begin
    SelectedRoot := NormalizeDriveRoot(WizardForm.DirEdit.Text);
    if not IsAllowedInstallDrive(SelectedRoot) then
    begin
      MsgBox(
        'Instalaciyata e pozvolena samo vurhu USB flashka, USB HDD ili USB SSD.' + #13#10#13#10 +
        'Lokalna instalaciya na tozi kompyutur ne e pozvolena.',
        mbError,
        MB_OK
      );
      Result := False;
    end;
  end;
end;

function GetInstalledAppExe: string;
begin
  Result := ExpandConstant('{app}\{#MyAppExeName}');
end;

function GetInstalledIconPath: string;
begin
  Result := ExpandConstant('{app}\_internal\assets\wga-icon.ico');
end;

function GetInstallDriveRoot: string;
begin
  Result := NormalizeDriveRoot(ExpandConstant('{app}'));
end;

procedure HideUsbRootFile(const FileName: string);
var
  ResultCode: Integer;
begin
  if not FileExists(FileName) then
    Exit;

  Exec(
    ExpandConstant('{cmd}'),
    '/C attrib +h +s "' + FileName + '"',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode
  );
end;

procedure CreateUsbLauncherFiles;
var
  DriveRoot: string;
  IconTarget: string;
  VbsTarget: string;
  CmdTarget: string;
  BatTarget: string;
  ReadmeTarget: string;
  AutorunTarget: string;
  AutorunText: string;
  VbsText: string;
  CmdText: string;
  ReadmeText: string;
begin
  DriveRoot := GetInstallDriveRoot();
  if DriveRoot = '' then
    Exit;

  IconTarget := DriveRoot + 'WGA-Drive.ico';
  VbsTarget := DriveRoot + 'WGA-Launch.vbs';
  CmdTarget := DriveRoot + 'Start WGA.cmd';
  BatTarget := DriveRoot + 'start.bat';
  ReadmeTarget := DriveRoot + 'WGA-USB-README.txt';
  AutorunTarget := DriveRoot + 'autorun.inf';

  if FileExists(GetInstalledIconPath()) then
    CopyFile(GetInstalledIconPath(), IconTarget, False);

  VbsText :=
    'Set oShell = CreateObject("WScript.Shell")' + #13#10 +
    'Set fso = CreateObject("Scripting.FileSystemObject")' + #13#10 +
    'root = fso.GetParentFolderName(WScript.ScriptFullName)' + #13#10 +
    'If Right(root, 1) <> "\" Then root = root & "\"' + #13#10 +
    'oShell.Run Chr(34) & root & "Start WGA.cmd" & Chr(34), 1, False' + #13#10;
  SaveStringToFile(VbsTarget, VbsText, False);

  CmdText :=
    '@echo off' + #13#10 +
    'setlocal' + #13#10 +
    'set "ROOT=%~dp0"' + #13#10 +
    'set "APP_EXE="' + #13#10 +
    'if exist "%ROOT%WinSys Guardian Advanced\WGA.exe" set "APP_EXE=%ROOT%WinSys Guardian Advanced\WGA.exe"' + #13#10 +
    'if not defined APP_EXE if exist "%ROOT%WGA\WGA.exe" set "APP_EXE=%ROOT%WGA\WGA.exe"' + #13#10 +
    'if not defined APP_EXE if exist "%ROOT%WGA.exe" set "APP_EXE=%ROOT%WGA.exe"' + #13#10 +
    'if not defined APP_EXE (' + #13#10 +
    '  for /f "delims=" %%F in (''dir /b /s "%ROOT%WGA.exe" 2^>nul'') do (' + #13#10 +
    '    set "APP_EXE=%%F"' + #13#10 +
    '    goto :FOUND_WGA' + #13#10 +
    '  )' + #13#10 +
    ')' + #13#10 +
    ':FOUND_WGA' + #13#10 +
    'if not defined APP_EXE (' + #13#10 +
    '  echo WGA.exe was not found on this drive.' + #13#10 +
    '  echo Start this file from the USB flash/HDD/SSD where WGA was installed.' + #13#10 +
    '  pause' + #13#10 +
    '  exit /b 1' + #13#10 +
    ')' + #13#10 +
    'start "" "%APP_EXE%"' + #13#10;
  SaveStringToFile(CmdTarget, CmdText, False);
  SaveStringToFile(BatTarget, CmdText, False);

  ReadmeText :=
    'WinSys Guardian Advanced USB package' + #13#10#13#10 +
    '1. Modern Windows systems may block autorun from USB devices.' + #13#10 +
    '2. If autorun does not start, open "Start WGA.cmd" from the root of the device.' + #13#10 +
    '3. You can also open "start.bat"; it searches this drive for the installed WGA.exe.' + #13#10 +
    '4. Installed application path: ' + ExpandConstant('{app}') + #13#10;
  SaveStringToFile(ReadmeTarget, ReadmeText, False);

  AutorunText :=
    '[autorun]' + #13#10 +
    'open=WGA-Launch.vbs' + #13#10 +
    'action=Open WinSys Guardian Advanced' + #13#10 +
    'label=WinSys Guardian Advanced' + #13#10 +
    'icon=WGA-Drive.ico' + #13#10 +
    'useautoplay=1' + #13#10 +
    'shell\open=Open WinSys Guardian Advanced' + #13#10 +
    'shell\open\command=WGA-Launch.vbs' + #13#10;
  SaveStringToFile(AutorunTarget, AutorunText, False);

  HideUsbRootFile(IconTarget);
  HideUsbRootFile(VbsTarget);
  HideUsbRootFile(AutorunTarget);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    CreateUsbLauncherFiles;
end;

procedure InitializeWizard;
begin
  WizardForm.SelectDirLabel.Caption :=
    'Izberi papka za instalaciya samo vurhu USB flashka, USB HDD ili USB SSD. Lokalna instalaciya ne e pozvolena.';
  WizardForm.DirEdit.Text := GetSuggestedInstallDir('');
end;
