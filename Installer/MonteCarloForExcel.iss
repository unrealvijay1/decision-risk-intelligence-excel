#define MyAppName "Monte Carlo for Excel"
#define MyAppPublisher "Vijay"
#define MyAppFileName "MonteCarloForExcel.xll"
#ifndef ProductVersion
  #error ProductVersion must be supplied by build-installer.ps1
#endif
#ifndef PayloadDir
  #error PayloadDir must be a verified Release staging directory
#endif
#ifndef OutputDirectory
  #error OutputDirectory must be supplied
#endif
#define MyAppVersion ProductVersion
#define SourceXll PayloadDir + "\MonteCarlo.Excel-AddIn64-packed.xll"
#define SourceXll32 PayloadDir + "\MonteCarlo.Excel-AddIn-packed.xll"

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
OutputDir={#OutputDirectory}
OutputBaseFilename=MonteCarloForExcel-Setup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes
SetupMutex=MonteCarloForExcelSetup
CloseApplications=no
RestartApplications=no
MinVersion=10.0.19045
UninstallDisplayName={#MyAppName}
VersionInfoVersion={#MyAppVersion}.0
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription=Monte Carlo simulation add-in for Microsoft Excel
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}

[Files]
Source: "{#SourceXll}"; DestDir: "{app}"; DestName: "{#MyAppFileName}"; Flags: ignoreversion; Check: UseX64
Source: "{#SourceXll32}"; DestDir: "{app}"; DestName: "{#MyAppFileName}"; Flags: ignoreversion; Check: UseX86
Source: "{#SourceXll}.config"; DestDir: "{app}"; DestName: "{#MyAppFileName}.config"; Flags: onlyifdoesntexist uninsneveruninstall; Check: UseX64
Source: "{#SourceXll32}.config"; DestDir: "{app}"; DestName: "{#MyAppFileName}.config"; Flags: onlyifdoesntexist uninsneveruninstall; Check: UseX86
Source: "{#PayloadDir}\windowsdesktop-runtime-10.0.12-win-x64.exe"; Flags: dontcopy
Source: "{#PayloadDir}\windowsdesktop-runtime-10.0.12-win-x86.exe"; Flags: dontcopy
Source: "{#PayloadDir}\CustomerGuide.txt"; DestDir: "{app}"; Flags: ignoreversion

[Code]
type
  TRegistrySnapshot = record
    Key, Name, Data, Expected: String;
    Existed, Changed, Deleted: Boolean;
  end;
const
  InstallerRegistryKey = 'Software\MonteCarloExcel\Installer';
  DefaultOptionsKey = 'Software\Microsoft\Office\16.0\Excel\Options';
var
  ExcelPage: TInputFileWizardPage;
  ExcelPath, ExcelArchitecture, ExcelOptionsKey: String;
  Candidates: TStringList;
  RuntimeReboot: Boolean;
  PreviousOptionsKey, PreviousOpenCommand, PreviousDirectory: String;
  RegistrationSnapshot: array of TRegistrySnapshot;
  RegistrationReady, InstallationComplete: Boolean;

function RememberValue(Key, Name: String): Integer;
begin
  Result := GetArrayLength(RegistrationSnapshot);
  SetArrayLength(RegistrationSnapshot, Result+1);
  RegistrationSnapshot[Result].Key := Key;
  RegistrationSnapshot[Result].Name := Name;
  RegistrationSnapshot[Result].Existed := RegQueryStringValue(HKCU, Key, Name, RegistrationSnapshot[Result].Data);
end;
function OwnedWrite(Key, Name, Data: String): Boolean;
var Index: Integer;
begin
  Index := RememberValue(Key, Name);
  Result := RegWriteStringValue(HKCU, Key, Name, Data);
  RegistrationSnapshot[Index].Changed := Result;
  RegistrationSnapshot[Index].Expected := Data;
end;
function OwnedDelete(Key, Name: String): Boolean;
var Index: Integer;
begin
  Index := RememberValue(Key, Name);
  Result := RegDeleteValue(HKCU, Key, Name);
  RegistrationSnapshot[Index].Changed := Result;
  RegistrationSnapshot[Index].Deleted := True;
end;
procedure RollbackRegistration();
var Index: Integer; Data: String; Matches: Boolean;
begin
  for Index := GetArrayLength(RegistrationSnapshot)-1 downto 0 do
    if RegistrationSnapshot[Index].Changed then begin
      if RegistrationSnapshot[Index].Deleted then
        Matches := not RegValueExists(HKCU, RegistrationSnapshot[Index].Key, RegistrationSnapshot[Index].Name)
      else
        Matches := RegQueryStringValue(HKCU, RegistrationSnapshot[Index].Key, RegistrationSnapshot[Index].Name, Data) and
          (CompareText(Data, RegistrationSnapshot[Index].Expected) = 0);
      if Matches then begin
        if RegistrationSnapshot[Index].Existed then
          RegWriteStringValue(HKCU, RegistrationSnapshot[Index].Key, RegistrationSnapshot[Index].Name, RegistrationSnapshot[Index].Data)
        else RegDeleteValue(HKCU, RegistrationSnapshot[Index].Key, RegistrationSnapshot[Index].Name);
      end;
    end;
end;

function UseX64(): Boolean;
begin Result := ExcelArchitecture = 'x64'; end;
function UseX86(): Boolean;
begin Result := ExcelArchitecture = 'x86'; end;
function GetAddInPath(): String;
begin Result := ExpandConstant('{app}\{#MyAppFileName}'); end;
function CommandFor(Path: String): String;
begin Result := '/R "' + Path + '"'; end;
function GetExcelOpenCommand(): String;
begin Result := CommandFor(GetAddInPath()); end;
function GetOpenValueName(Index: Integer): String;
begin
  if Index = 0 then Result := 'OPEN' else Result := 'OPEN' + IntToStr(Index);
end;
function IsOpenName(Name: String): Boolean;
var N: Integer;
begin
  Result := CompareText(Name, 'OPEN') = 0;
  if not Result and (CompareText(Copy(Name, 1, 4), 'OPEN') = 0) then begin
    N := StrToIntDef(Copy(Name, 5, Length(Name)), -1);
    Result := (N > 0) and (CompareText(Name, GetOpenValueName(N)) = 0);
  end;
end;

function ReadArchitecture(Path: String): String;
var Bytes: AnsiString; Offset, Machine: Integer;
begin
  Result := '';
  if not LoadStringFromFile(Path, Bytes) or (Length(Bytes) < 64) then Exit;
  if Copy(Bytes, 1, 2) <> 'MZ' then Exit;
  Offset := Ord(Bytes[61]) + Ord(Bytes[62])*256 + Ord(Bytes[63])*65536 + Ord(Bytes[64])*16777216;
  if (Offset < 64) or (Offset > Length(Bytes)-6) then Exit;
  if Copy(Bytes, Offset+1, 4) <> 'PE'+#0+#0 then Exit;
  Machine := Ord(Bytes[Offset+5]) + Ord(Bytes[Offset+6])*256;
  if Machine = $014C then Result := 'x86';
  if Machine = $8664 then Result := 'x64';
end;
function SelectExcel(Path: String): String;
var MS, LS, Major: Cardinal;
begin
  Result := '';
  Path := RemoveQuotes(Trim(Path));
  if not FileExists(Path) or (CompareText(ExtractFileName(Path), 'EXCEL.EXE') <> 0) then begin
    Result := 'Select the installed Microsoft EXCEL.EXE.'; Exit;
  end;
  ExcelArchitecture := ReadArchitecture(Path);
  if ExcelArchitecture = '' then begin
    Result := 'Only x86 and x64 desktop Excel are supported. The selected file has an unsupported PE architecture.'; Exit;
  end;
  if not GetVersionNumbers(Path, MS, LS) then begin
    Result := 'Cannot read the selected Excel version.'; Exit;
  end;
  Major := MS shr 16;
  if Major < 14 then begin Result := 'Excel 2010 or newer is required.'; Exit; end;
  ExcelOptionsKey := 'Software\Microsoft\Office\' + IntToStr(Major) + '.0\Excel\Options';
  ExcelPath := Path;
end;
procedure AddCandidate(Path: String);
begin
  Path := RemoveQuotes(Trim(Path));
  if FileExists(Path) and (Candidates.IndexOf(Path) < 0) then Candidates.Add(Path);
end;
procedure Discover(Root: Integer);
var Path: String; Major: Integer;
begin
  if RegQueryStringValue(Root, 'Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', '', Path) then AddCandidate(Path);
  for Major := 14 to 16 do
    if RegQueryStringValue(Root, 'Software\Microsoft\Office\'+IntToStr(Major)+'.0\Excel\InstallRoot', 'Path', Path) then
      AddCandidate(AddBackslash(Path)+'EXCEL.EXE');
  if RegQueryStringValue(Root, 'Software\Microsoft\Office\ClickToRun\Configuration', 'InstallationPath', Path) then
    AddCandidate(AddBackslash(Path)+'root\Office16\EXCEL.EXE');
end;
procedure InitializeWizard();
var Path, Error: String;
begin
  Candidates := TStringList.Create;
  PreviousOptionsKey := DefaultOptionsKey;
  RegQueryStringValue(HKCU, InstallerRegistryKey, 'ExcelOptionsKey', PreviousOptionsKey);
  PreviousOpenCommand := '';
  RegQueryStringValue(HKCU, InstallerRegistryKey, 'ExcelOpenCommand', PreviousOpenCommand);
  PreviousDirectory := '';
  RegQueryStringValue(HKCU, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{9B4D8E57-7B33-4C68-9D1D-91D5E4D8F001}_is1', 'InstallLocation', PreviousDirectory);
  if (PreviousOpenCommand = '') and (PreviousDirectory <> '') then
    PreviousOpenCommand := CommandFor(AddBackslash(PreviousDirectory)+'{#MyAppFileName}');
  Discover(HKCU32); Discover(HKLM32);
  if IsWin64 then begin Discover(HKCU64); Discover(HKLM64); end;
  Path := ExpandConstant('{param:EXCELPATH|}');
  if (Path = '') and (Candidates.Count = 1) then Path := Candidates[0];
  ExcelPage := CreateInputFilePage(wpWelcome, 'Select Microsoft Excel',
    'Setup installs the XLL that matches Excel, not Windows.',
    'Select the desktop EXCEL.EXE you use. Ambiguous detection requires an explicit choice.');
  ExcelPage.Add('Excel executable:', 'Microsoft Excel|EXCEL.EXE', '.exe');
  ExcelPage.Values[0] := Path;
  if Path <> '' then begin
    Error := SelectExcel(Path);
    if Error <> '' then Log(Error);
  end;
  if WizardSilent and (ExcelPath = '') then begin
    Log('Excel detection is inconclusive. Supply /EXCELPATH="<full path>".');
    SuppressibleMsgBox('Excel detection is inconclusive. Supply /EXCELPATH="<full path>".', mbError, MB_OK, IDOK);
    Abort;
  end;
end;
function ShouldSkipPage(PageID: Integer): Boolean;
begin Result := (PageID = ExcelPage.ID) and (ExcelPath <> ''); end;
function NextButtonClick(CurPageID: Integer): Boolean;
var Error: String;
begin
  Result := True;
  if CurPageID = ExcelPage.ID then begin
    Error := SelectExcel(ExcelPage.Values[0]);
    Result := Error = '';
    if not Result then SuppressibleMsgBox(Error, mbError, MB_OK, IDOK);
  end;
end;
function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo, MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
begin
  Result := MemoDirInfo + NewLine + NewLine + 'Excel: ' + ExcelPath + NewLine +
    'Architecture: ' + ExcelArchitecture + NewLine + 'Version: {#MyAppVersion}' + NewLine +
    'Your workbooks and logging configuration are preserved. No activation is required.';
end;
function ExcelRunning(): Boolean;
var Locator, Service, Processes: Variant;
begin
  { Fail closed if process discovery cannot be completed. Never terminate Excel. }
  Result := True;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    Service := Locator.ConnectServer('', 'root\CIMV2');
    Processes := Service.ExecQuery('SELECT ProcessId FROM Win32_Process WHERE Name="EXCEL.EXE"');
    Result := Processes.Count > 0;
  except
    Log('Cannot verify Excel process state. Close Excel and retry with WMI available.');
  end;
end;
function RuntimeInstalled(): Boolean;
var Root, I, Patch: Integer; Names: TArrayOfString; Key: String;
begin
  Result := False;
  { Microsoft registers all architecture-qualified .NET setup keys in the
    32-bit registry view, including x64 runtimes on x64 Windows. }
  Root := HKLM32;
  Key := 'SOFTWARE\dotnet\Setup\InstalledVersions\'+ExcelArchitecture+'\sharedfx\Microsoft.WindowsDesktop.App';
  if RegGetValueNames(Root, Key, Names) then
    for I := 0 to GetArrayLength(Names)-1 do
      if Copy(Names[I], 1, 5) = '10.0.' then begin
        Patch := StrToIntDef(Copy(Names[I], 6, Length(Names[I])), -1);
        if Patch >= 12 then Result := True;
      end;
end;
function FindRegistrationSlot(): String;
var I: Integer; Data, Name: String;
begin
  Result := '';
  for I := 0 to 4095 do begin
    Name := GetOpenValueName(I);
    if RegQueryStringValue(HKCU, ExcelOptionsKey, Name, Data) and
       (CompareText(Data, GetExcelOpenCommand()) = 0) then begin Result := Name; Exit; end;
  end;
  for I := 0 to 4095 do begin
    Name := GetOpenValueName(I);
    if not RegValueExists(HKCU, ExcelOptionsKey, Name) then begin Result := Name; Exit; end;
  end;
end;
procedure RegisterExcelAddIn(); forward;
function PrepareToInstall(var NeedsRestart: Boolean): String;
var RuntimeFile: String; Code: Integer;
begin
  Result := '';
  if ExcelPath = '' then begin
    Result := 'Excel detection is inconclusive. Select EXCEL.EXE or supply /EXCELPATH="<full path>".'; Exit;
  end;
  if ExcelRunning then begin Result := 'Close all Microsoft Excel instances before installing or updating Monte Carlo for Excel.'; Exit; end;
  if FindRegistrationSlot() = '' then begin Result := 'No available Excel OPEN registry slot. Installation cancelled.'; Exit; end;
  if not RuntimeInstalled then begin
    if WizardSilent and not IsAdmin then begin
      Result := 'Microsoft .NET 10 Desktop Runtime ('+ExcelArchitecture+', 10.0.12 or newer) is missing. Install it with administrator approval, then retry.'; Exit;
    end;
    RuntimeFile := 'windowsdesktop-runtime-10.0.12-win-'+ExcelArchitecture+'.exe';
    ExtractTemporaryFile(RuntimeFile);
    if not ShellExec('runas', ExpandConstant('{tmp}\')+RuntimeFile, '/install /passive /norestart', '', SW_SHOWNORMAL, ewWaitUntilTerminated, Code) then begin
      Result := 'Desktop Runtime installation was cancelled or could not start.'; Exit;
    end;
    RuntimeReboot := Code = 3010;
    if (Code <> 0) and not RuntimeReboot then
      Result := 'Desktop Runtime installation failed. Exit code: '+IntToStr(Code)+'. Retry after installing the matching .NET 10 Desktop Runtime.'
    else if not RuntimeInstalled then
      Result := 'The runtime installer completed successfully, but the matching '+ExcelArchitecture+' Desktop Runtime registration could not be verified. Restart Windows if required, then retry Setup. See the Setup log.';
    NeedsRestart := RuntimeReboot;
  end;
  if (Result = '') and not RegistrationReady then begin
    try
      RegisterExcelAddIn();
      RegistrationReady := True;
    except
      Result := 'Could not register Monte Carlo for Excel. '+GetExceptionMessage();
    end;
  end;
end;
procedure DeleteOwnedRegistrations(Key, Command, ExceptName: String);
var Names: TArrayOfString; I: Integer; Data: String;
begin
  if (Command = '') or not RegGetValueNames(HKCU, Key, Names) then Exit;
  for I := 0 to GetArrayLength(Names)-1 do
    if IsOpenName(Names[I]) and (CompareText(Names[I], ExceptName) <> 0) and
       RegQueryStringValue(HKCU, Key, Names[I], Data) and (CompareText(Data, Command) = 0) then
      if not OwnedDelete(Key, Names[I]) then RaiseException('Could not remove owned duplicate Excel registration.');
end;
procedure RegisterExcelAddIn();
var Name, PreviousKey, PreviousCommand: String;
begin
  SetArrayLength(RegistrationSnapshot, 0);
  try
  { Snapshot taken before Inno updates its uninstall InstallLocation. }
  PreviousKey := PreviousOptionsKey;
  PreviousCommand := PreviousOpenCommand;
  Name := FindRegistrationSlot();
  if Name = '' then RaiseException('No free Excel registration slot.');
  if not OwnedWrite(ExcelOptionsKey, Name, GetExcelOpenCommand()) then
    RaiseException('Could not write Excel add-in registration.');
  if not OwnedWrite(InstallerRegistryKey, 'ExcelOpenValue', Name) or
     not OwnedWrite(InstallerRegistryKey, 'ExcelOpenCommand', GetExcelOpenCommand()) or
     not OwnedWrite(InstallerRegistryKey, 'ExcelOptionsKey', ExcelOptionsKey) or
     not OwnedWrite(InstallerRegistryKey, 'ExcelArchitecture', ExcelArchitecture) or
     not OwnedWrite(InstallerRegistryKey, 'ExcelPath', ExcelPath) or
     not OwnedWrite(InstallerRegistryKey, 'ProductVersion', '{#MyAppVersion}') then begin
    RaiseException('Could not save installer registration ownership.');
  end;
  DeleteOwnedRegistrations(ExcelOptionsKey, GetExcelOpenCommand(), Name);
  if CompareText(PreviousKey, ExcelOptionsKey) = 0 then
    DeleteOwnedRegistrations(PreviousKey, PreviousCommand, Name)
  else DeleteOwnedRegistrations(PreviousKey, PreviousCommand, '');
  except
    RollbackRegistration();
    RaiseException(GetExceptionMessage());
  end;
end;
function InitializeUninstall(): Boolean;
begin
  Result := not ExcelRunning;
  if not Result then SuppressibleMsgBox('Close all Microsoft Excel instances before uninstalling Monte Carlo for Excel.', mbError, MB_OK, IDOK);
end;
procedure UnregisterExcelAddIn();
var Key, Command, Recorded: String;
begin
  Key := DefaultOptionsKey;
  RegQueryStringValue(HKCU, InstallerRegistryKey, 'ExcelOptionsKey', Key);
  Command := GetExcelOpenCommand();
  { Only exact installed-path entries are owned; never erase a replaced entry. }
  DeleteOwnedRegistrations(Key, Command, '');
  Recorded := '';
  RegQueryStringValue(HKCU, InstallerRegistryKey, 'ExcelOpenCommand', Recorded);
  if CompareText(Recorded, Command) = 0 then begin
    RegDeleteValue(HKCU, InstallerRegistryKey, 'ExcelOpenValue');
    RegDeleteValue(HKCU, InstallerRegistryKey, 'ExcelOpenCommand');
    RegDeleteValue(HKCU, InstallerRegistryKey, 'ExcelOptionsKey');
    RegDeleteValue(HKCU, InstallerRegistryKey, 'ExcelArchitecture');
    RegDeleteValue(HKCU, InstallerRegistryKey, 'ExcelPath');
    RegDeleteValue(HKCU, InstallerRegistryKey, 'ProductVersion');
    RegDeleteKeyIfEmpty(HKCU, InstallerRegistryKey);
  end;
end;
procedure CurStepChanged(CurStep: TSetupStep);
var PreviousConfig, CurrentConfig, PreviousXll: String;
begin
  if (CurStep = ssInstall) and (PreviousDirectory <> '') and
     (CompareText(RemoveBackslashUnlessRoot(PreviousDirectory), ExpandConstant('{app}')) <> 0) then begin
    PreviousConfig := AddBackslash(PreviousDirectory)+'{#MyAppFileName}.config';
    CurrentConfig := GetAddInPath()+'.config';
    if FileExists(PreviousConfig) and not FileExists(CurrentConfig) then begin
      if not ForceDirectories(ExpandConstant('{app}')) or not CopyFile(PreviousConfig, CurrentConfig, True) then
        RaiseException('Could not preserve the previous logging configuration.');
    end;
  end;
  if CurStep = ssPostInstall then begin
    InstallationComplete := True;
    if PreviousDirectory <> '' then begin
      PreviousXll := AddBackslash(PreviousDirectory)+'{#MyAppFileName}';
      if (CompareText(PreviousXll, GetAddInPath()) <> 0) and
         (CompareText(PreviousOpenCommand, CommandFor(PreviousXll)) = 0) and FileExists(PreviousXll) then
        if not DeleteFile(PreviousXll) then Log('Could not remove obsolete XLL: '+PreviousXll);
    end;
  end;
end;
procedure DeinitializeSetup();
begin
  if RegistrationReady and not InstallationComplete then RollbackRegistration();
end;
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin if CurUninstallStep = usUninstall then UnregisterExcelAddIn(); end;
