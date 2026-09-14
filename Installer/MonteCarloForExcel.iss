#define MyAppName "Monte Carlo for Excel"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "Vijay"
#define MyAppFileName "MonteCarloForExcel.xll"

#define SourceXll "C:\montecarlo\MonteCarlo.Excel\MonteCarlo.Excel\bin\Release\net10.0-windows\publish\MonteCarlo.Excel-AddIn64-packed.xll"


[Setup]

AppId={{9B4D8E57-7B33-4C68-9D1D-91D5E4D8F001}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}

DefaultDirName={localappdata}\Programs\Monte Carlo for Excel

PrivilegesRequired=lowest

DisableProgramGroupPage=yes
DisableDirPage=yes

ArchitecturesAllowed=x64compatible

OutputDir=C:\montecarlo\Installer\Output
OutputBaseFilename=MonteCarloForExcel-Setup-{#MyAppVersion}

Compression=lzma2
SolidCompression=yes

WizardStyle=modern

SetupLogging=yes

UninstallDisplayName={#MyAppName}
Uninstallable=yes

VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription=Monte Carlo simulation add-in for Microsoft Excel
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}


[Files]

Source: "{#SourceXll}"; DestDir: "{app}"; DestName: "{#MyAppFileName}"; Flags: ignoreversion


[Code]

const
  ExcelOptionsKey = 'Software\Microsoft\Office\16.0\Excel\Options';
  InstallerRegistryKey = 'Software\MonteCarloExcel\Installer';
  InstallerOpenValueName = 'ExcelOpenValue';


function GetAddInPath(): String;
begin
  Result := ExpandConstant('{app}\{#MyAppFileName}');
end;


function GetExcelOpenCommand(): String;
begin
  Result := '/R "' + GetAddInPath() + '"';
end;


function GetOpenValueName(Index: Integer): String;
begin
  if Index = 0 then
  begin
    Result := 'OPEN';
  end
  else
  begin
    Result := 'OPEN' + IntToStr(Index);
  end;
end;


function FindExistingRegistration(
  var ExistingValueName: String): Boolean;

var
  Index: Integer;
  ValueName: String;
  ExistingData: String;

begin
  Result := False;
  ExistingValueName := '';

  for Index := 0 to 200 do
  begin
    ValueName := GetOpenValueName(Index);

    if RegQueryStringValue(
         HKCU,
         ExcelOptionsKey,
         ValueName,
         ExistingData) then
    begin
      if CompareText(
           ExistingData,
           GetExcelOpenCommand()) = 0 then
      begin
        ExistingValueName := ValueName;
        Result := True;
        Exit;
      end;
    end;
  end;
end;


function FindFreeOpenValue(): String;

var
  Index: Integer;
  ValueName: String;

begin
  Result := '';

  for Index := 0 to 200 do
  begin
    ValueName := GetOpenValueName(Index);

    if not RegValueExists(
             HKCU,
             ExcelOptionsKey,
             ValueName) then
    begin
      Result := ValueName;
      Exit;
    end;
  end;
end;


procedure RegisterExcelAddIn();

var
  ValueName: String;
  ExistingValueName: String;
  OpenCommand: String;

begin
  OpenCommand := GetExcelOpenCommand();

  if FindExistingRegistration(ExistingValueName) then
  begin
    RegWriteStringValue(
      HKCU,
      InstallerRegistryKey,
      InstallerOpenValueName,
      ExistingValueName);

    Exit;
  end;

  ValueName := FindFreeOpenValue();

  if ValueName = '' then
  begin
  MsgBox(
    'Monte Carlo for Excel was installed, but the installer ' +
    'could not find an available Excel OPEN registry entry. ' +
    'The add-in can still be loaded manually through Excel.',
    mbError,
    MB_OK);

    Exit;
  end;

  if not RegWriteStringValue(
           HKCU,
           ExcelOptionsKey,
           ValueName,
           OpenCommand) then
  begin
    MsgBox(
      'The installer could not register Monte Carlo for Excel ' +
      'with Microsoft Excel.',
      mbError,
      MB_OK);

    Exit;
  end;

  RegWriteStringValue(
    HKCU,
    InstallerRegistryKey,
    InstallerOpenValueName,
    ValueName);
end;


procedure UnregisterExcelAddIn();

var
  ValueName: String;
  ExistingData: String;

begin
  if not RegQueryStringValue(
           HKCU,
           InstallerRegistryKey,
           InstallerOpenValueName,
           ValueName) then
  begin
    Exit;
  end;

  if RegQueryStringValue(
       HKCU,
       ExcelOptionsKey,
       ValueName,
       ExistingData) then
  begin
    if CompareText(
         ExistingData,
         GetExcelOpenCommand()) = 0 then
    begin
      RegDeleteValue(
        HKCU,
        ExcelOptionsKey,
        ValueName);
    end;
  end;

  RegDeleteValue(
    HKCU,
    InstallerRegistryKey,
    InstallerOpenValueName);
end;


procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    RegisterExcelAddIn();
  end;
end;


procedure CurUninstallStepChanged(
  CurUninstallStep: TUninstallStep);

begin
  if CurUninstallStep = usUninstall then
  begin
    UnregisterExcelAddIn();
  end;
end;