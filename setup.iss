; Inno Setup script for "Tools By Sanat". Build with build-setup.ps1 (passes /DMyAppVersion=x.y.z).
; Per-user install, no administrator rights needed.

#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif
#define MyAppName "Tools By Sanat"

[Setup]
AppId={{B1D0C7F2-5E3A-4C8B-9A64-2F7E1D3C8A51}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=Sanat
; {app} only holds the uninstaller and register.ps1; the add-in files go to Excel's AddIns folder (see [Files])
DefaultDirName={localappdata}\Programs\ToolsBySanat
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=dist
OutputBaseFilename=ToolsBySanat-Setup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#MyAppName}
CloseApplications=yes
RestartApplications=no

[Files]
Source: "build\ToolsBySanat.xlam";       DestDir: "{userappdata}\Microsoft\AddIns"; Flags: ignoreversion
Source: "build\ToolsBySanatLoader.xlam"; DestDir: "{userappdata}\Microsoft\AddIns"; Flags: ignoreversion
Source: "register.ps1";                  DestDir: "{app}";                          Flags: ignoreversion

[InstallDelete]
; a download that was waiting for the next start is older than what is being installed
Type: filesandordirs; Name: "{userappdata}\Microsoft\AddIns\ToolsBySanat_data\pending"

[UninstallDelete]
Type: filesandordirs; Name: "{userappdata}\Microsoft\AddIns\ToolsBySanat_data\pending"
Type: filesandordirs; Name: "{userappdata}\Microsoft\AddIns\ToolsBySanat_data\prev"

[Code]
function IsExcelRunning: Boolean;
begin
  Result := FindWindowByClassName('XLMAIN') <> 0;
end;

function RunRegister(const Action: String): Integer;
var
  ResultCode: Integer;
begin
  if not Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
      '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + ExpandConstant('{app}\register.ps1') + '" -Action ' + Action,
      '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    ResultCode := 1;
  Result := ResultCode;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  while IsExcelRunning do
  begin
    if SuppressibleMsgBox('Please close all Excel windows, then click Retry.' + #13#10 +
                          'Excel keeps the add-in files locked while it is open.',
                          mbInformation, MB_RETRYCANCEL, IDCANCEL) = IDCANCEL then
    begin
      Result := 'Setup cannot continue while Excel is open. Close Excel and run Setup again.';
      Exit;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  Code: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    Code := RunRegister('Install');
    if Code = 2 then
      SuppressibleMsgBox('Microsoft Excel was not found, so the add-in could not be registered.' + #13#10 +
                         'Install Excel and run Setup again.', mbError, MB_OK, IDOK)
    else if Code <> 0 then
      SuppressibleMsgBox('The add-in was copied, but could not be registered in Excel automatically.' + #13#10 +
                         'In Excel: File > Options > Add-ins > Go... > Browse, and select ToolsBySanatLoader.xlam in:' + #13#10 +
                         ExpandConstant('{userappdata}\Microsoft\AddIns'), mbError, MB_OK, IDOK);
  end;
end;

function InitializeUninstall: Boolean;
begin
  Result := True;
  while IsExcelRunning do
  begin
    if SuppressibleMsgBox('Please close all Excel windows, then click Retry.', mbInformation, MB_RETRYCANCEL, IDCANCEL) = IDCANCEL then
    begin
      Result := False;
      Exit;
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  DataDir: String;
begin
  if CurUninstallStep = usUninstall then
    RunRegister('Uninstall')     // runs before the files are removed
  else if CurUninstallStep = usPostUninstall then
  begin
    DataDir := ExpandConstant('{userappdata}\Microsoft\AddIns\ToolsBySanat_data');
    if SuppressibleMsgBox('Also delete your settings files (settings.csv and party_master.csv)?' + #13#10 +
                          'Choose No to keep them.', mbConfirmation, MB_YESNO or MB_DEFBUTTON2, IDNO) = IDYES then
    begin
      DeleteFile(DataDir + '\settings.csv');
      DeleteFile(DataDir + '\party_master.csv');
    end;
    RemoveDir(DataDir);          // only removes it if empty
  end;
end;
