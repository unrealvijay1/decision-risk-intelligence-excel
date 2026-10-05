[CmdletBinding()]
param([Parameter(Mandatory)][string]$PayloadDir,[string]$IsccPath='C:\Program Files\Inno Setup 7\ISCC.exe',[string]$Version='0.2.0',[switch]$LiveStartup)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot
$runId=[guid]::NewGuid().ToString('N')
$root=Join-Path $repo ('artifacts/installer-tests-'+$runId)
[void](New-Item -ItemType Directory -Force $root)
$namespace='Software\MonteCarloExcel\InstallerTests\'+$runId
$testAppId=[guid]::NewGuid().ToString().ToUpper()
$original=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'MonteCarloForExcel.iss'))
$source=$original.Replace('9B4D8E57-7B33-4C68-9D1D-91D5E4D8F001',$testAppId)
$source=$source.Replace("InstallerRegistryKey = 'Software\MonteCarloExcel\Installer'",("InstallerRegistryKey = '"+$namespace+"\Installer'"))
$source=$source.Replace("DefaultOptionsKey = 'Software\Microsoft\Office\16.0\Excel\Options'",("DefaultOptionsKey = '"+$namespace+"\Office\16.0\Excel\Options'"))
$source=$source.Replace("ExcelOptionsKey := 'Software\Microsoft\Office\'",("ExcelOptionsKey := '"+$namespace+"\Office\'"))
$source=$source.Replace('Discover(HKCU32); Discover(HKLM32);','').Replace('if IsWin64 then begin Discover(HKCU64); Discover(HKLM64); end;','')
$source=$source.Replace('Major := MS shr 16;','Major := 16; { Test PE fixtures have product version 0.2, not an Office version. }')
$source=[regex]::Replace($source,'function ExcelRunning\(\): Boolean;[\s\S]*?(?=function RuntimeInstalled)',"function ExcelRunning(): Boolean;`nbegin Result := False; end;`n")
$source=[regex]::Replace($source,'function RuntimeInstalled\(\): Boolean;[\s\S]*?(?=function FindRegistrationSlot)',"function RuntimeInstalled(): Boolean;`nbegin Result := ExpandConstant('{param:RUNTIME|1}') <> '0'; end;`n")
# Missing-prerequisite tests must never invoke the real runtime installer on this host.
$source=[regex]::Replace($source,"if not ShellExec\('runas', [^\r\n]+ then begin",'if True then begin')
$source=$source.Replace('Index := RememberValue(Key, Name);',"if (ExpandConstant('{param:FAILREG|0}') = '1') and (Name = 'ExcelOpenCommand') then begin Result := False; Exit; end;`n  Index := RememberValue(Key, Name);")
$source=$source.Replace('Compression=lzma2','Compression=none').Replace('SolidCompression=yes','SolidCompression=no')
$source=$source.Replace('MonteCarloForExcelSetup','MonteCarloInstallerFixture'+$runId)
$script=Join-Path $root 'Fixture.iss'; [IO.File]::WriteAllText($script,$source)
& $IsccPath ('/DProductVersion='+$Version) ('/DPayloadDir='+$PayloadDir) ('/DOutputDirectory='+$root) $script | Set-Content (Join-Path $root 'compile.log')
if($LASTEXITCODE -ne 0){throw 'Fixture compiler failed'}
$setup=Join-Path $root "MonteCarloForExcel-Setup-$Version.exe"
$options=$namespace+'\Office\16.0\Excel\Options'; $meta=$namespace+'\Installer'
$arp='Software\Microsoft\Windows\CurrentVersion\Uninstall\{'+$testAppId+'}_is1'
$oldDir=Join-Path $root 'old-app'; $newDir=Join-Path $root 'new-app'
$exeDir=Join-Path $root 'excel'; [void](New-Item -ItemType Directory -Force $exeDir)
$excel=Join-Path $exeDir 'EXCEL.EXE'
Copy-Item -LiteralPath (Join-Path $repo 'Installer/Checks/bin/Release/net10.0-windows/Checks.exe') -Destination $excel
$count=0
function Assert([bool]$Condition,[string]$Name){if(-not $Condition){throw "FAIL $Name"}; $script:count++; Write-Output "PASS $Name"}
function SetValue([string]$Key,[string]$Name,[string]$Data){$k=[Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($Key);try{$k.SetValue($Name,$Data)}finally{$k.Dispose()}}
function ReadValue([string]$Key,[string]$Name){$k=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($Key);try{if($k){return $k.GetValue($Name)}}finally{if($k){$k.Dispose()}}}
function Values([string]$Key){$k=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($Key);try{if($k){foreach($name in $k.GetValueNames()){[pscustomobject]@{Name=$name;Data=$k.GetValue($name)}}}}finally{if($k){$k.Dispose()}}}
function Install([string]$Dir,[string]$Label){
  $p=Start-Process -FilePath $setup -ArgumentList @('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART',('/DIR="'+$Dir+'"'),('/EXCELPATH="'+$excel+'"'),('/LOG="'+(Join-Path $root ($Label+'.log'))+'"')) -WindowStyle Hidden -PassThru
  if(-not $p.WaitForExit(60000)){throw 'Fixture installation timed out; process retained for diagnosis.'}
  Assert ($p.ExitCode -eq 0) "$Label exits successfully"
}
function Uninstall([string]$Dir,[string]$Label){
  $p=Start-Process -FilePath (Join-Path $Dir 'unins000.exe') -ArgumentList @('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART',('/LOG="'+(Join-Path $root ($Label+'.log'))+'"')) -WindowStyle Hidden -PassThru
  if(-not $p.WaitForExit(60000)){throw 'Fixture uninstall timed out; process retained for diagnosis.'}
  Assert ($p.ExitCode -eq 0) "$Label exits successfully"
}
function ExpectFailure([string]$Label,[string[]]$Extra) {
  $dir=Join-Path $root $Label
  $log=Join-Path $root ($Label+'.log')
  $arguments=@('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART',('/DIR="'+$dir+'"'),('/LOG="'+$log+'"'))+$Extra
  $p=Start-Process -FilePath $setup -ArgumentList $arguments -WindowStyle Hidden -PassThru
  if(-not $p.WaitForExit(60000)){throw 'Fixture failure test timed out; process retained for diagnosis.'}
  Assert ($p.ExitCode -ne 0) "$Label fails clearly"
  return $log
}
try {
  SetValue $options OPEN '/R "C:\OtherVendor\Other.xll"'
  $protectedLicense=Join-Path $root 'persistent-license.dat'; $protectedTrial=Join-Path $root 'persistent-trial.dat'
  [IO.File]::WriteAllText($protectedLicense,'license sentinel'); [IO.File]::WriteAllText($protectedTrial,'trial sentinel')
  SetValue $namespace TrialState 'trial registry sentinel'
  $null=ExpectFailure missing-excel @()
  Assert (-not (Test-Path -LiteralPath (Join-Path $root 'missing-excel/MonteCarloForExcel.xll'))) 'inconclusive detection makes no installation'
  $null=ExpectFailure missing-runtime @(('/EXCELPATH="'+$excel+'"'),'/RUNTIME=0')
  Assert (-not (Test-Path -LiteralPath (Join-Path $root 'missing-runtime/MonteCarloForExcel.xll'))) 'missing runtime/cancelled elevation makes no installation'
  Install $oldDir fresh-x64
  Assert ((Get-FileHash (Join-Path $oldDir 'MonteCarloForExcel.xll')).Hash -eq (Get-FileHash (Join-Path $PayloadDir 'MonteCarlo.Excel-AddIn64-packed.xll')).Hash) 'fresh x64 payload matches verified Release'
  Assert ((ReadValue $meta ExcelArchitecture) -eq 'x64') 'actual PE detects x64'
  Assert ((ReadValue $options OPEN) -eq '/R "C:\OtherVendor\Other.xll"') 'fresh install preserves other add-in'
  Assert (@(Values $options | Where-Object Data -eq ('/R "'+$oldDir+'\MonteCarloForExcel.xll"')).Count -eq 1) 'fresh registration exactly once'
  if($LiveStartup) {
    # Starts a fresh private COM Excel; never attaches to an existing Excel process.
    & (Join-Path $repo 'MonteCarlo.Excel.StartupChecks/Verify-Startup.ps1') -ProductionPath (Join-Path $oldDir 'MonteCarloForExcel.xll') -UseInstalledCopy -FreshProfile
    if(-not $?){throw 'Installed-copy Excel verification failed'}
  }
  [IO.File]::WriteAllText((Join-Path $oldDir 'MonteCarloForExcel.xll.config'),'<configuration><monteCarloDiagnostics verbose="true" /></configuration>')
  Install $oldDir reinstall
  Assert (@(Values $options | Where-Object Data -eq ('/R "'+$oldDir+'\MonteCarloForExcel.xll"')).Count -eq 1) 'reinstall does not duplicate registration'
  Assert ([IO.File]::ReadAllText((Join-Path $oldDir 'MonteCarloForExcel.xll.config')).Contains('verbose="true"')) 'reinstall preserves explicit logging preference'
  $null=ExpectFailure registration-write-failure @(('/EXCELPATH="'+$excel+'"'),'/FAILREG=1')
  Assert (@(Values $options | Where-Object Data -eq ('/R "'+$oldDir+'\MonteCarloForExcel.xll"')).Count -eq 1) 'registration failure restores previous OPEN ownership'
  Assert ((ReadValue $meta ExcelOpenCommand) -eq ('/R "'+$oldDir+'\MonteCarloForExcel.xll"')) 'registration failure restores previous metadata'
  # Simulate legacy 0.1.0 installer ownership: only slot + stable AppId InstallLocation.
  [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($meta,$false)
  SetValue $meta ExcelOpenValue OPEN1
  SetValue $options OPEN2 ('/R "'+$oldDir+'\MonteCarloForExcel.xll"')
  Install $newDir upgrade-path
  Assert (@(Values $options | Where-Object Data -eq ('/R "'+$oldDir+'\MonteCarloForExcel.xll"')).Count -eq 0) 'legacy upgrade removes owned old path and duplicates'
  Assert (@(Values $options | Where-Object Data -eq ('/R "'+$newDir+'\MonteCarloForExcel.xll"')).Count -eq 1) 'upgrade registers new path exactly once'
  Assert ([IO.File]::ReadAllText((Join-Path $newDir 'MonteCarloForExcel.xll.config')).Contains('verbose="true"')) 'path migration preserves logging preference'
  Assert (-not (Test-Path -LiteralPath (Join-Path $oldDir 'MonteCarloForExcel.xll'))) 'path migration removes obsolete owned XLL'
  Assert ((ReadValue $options OPEN) -eq '/R "C:\OtherVendor\Other.xll"') 'upgrade preserves other add-in'
  Uninstall $newDir uninstall
  Assert (-not (Test-Path -LiteralPath (Join-Path $newDir 'MonteCarloForExcel.xll'))) 'uninstall removes XLL'
  Assert (Test-Path -LiteralPath (Join-Path $newDir 'MonteCarloForExcel.xll.config')) 'uninstall retains logging configuration'
  Assert (@(Values $options | Where-Object {$_.Data -like '*MonteCarloForExcel.xll*'}).Count -eq 0) 'uninstall removes owned registrations'
  Assert ((ReadValue $options OPEN) -eq '/R "C:\OtherVendor\Other.xll"') 'uninstall preserves other add-in'
  # Patch only the fixture PE header; file is never executed. Validates x86 Excel on x64 Windows.
  $bytes=[IO.File]::ReadAllBytes($excel); $offset=[BitConverter]::ToInt32($bytes,0x3c); $bytes[$offset+4]=0x4c; $bytes[$offset+5]=0x01; [IO.File]::WriteAllBytes($excel,$bytes)
  Install $newDir reinstall-x86
  Assert ((ReadValue $meta ExcelArchitecture) -eq 'x86') 'actual PE detects x86 on x64 Windows'
  Assert ((Get-FileHash (Join-Path $newDir 'MonteCarloForExcel.xll')).Hash -eq (Get-FileHash (Join-Path $PayloadDir 'MonteCarlo.Excel-AddIn-packed.xll')).Hash) 'x86 payload matches verified Release'
  $slot=ReadValue $meta ExcelOpenValue
  SetValue $options $slot '/R "C:\AnotherVendor\replacement.xll"'
  Uninstall $newDir uninstall-replaced-slot
  Assert ((ReadValue $options $slot) -eq '/R "C:\AnotherVendor\replacement.xll"') 'uninstall preserves replaced slot ownership'
  Assert ([IO.File]::ReadAllText($protectedLicense) -eq 'license sentinel') 'installer lifecycle preserves license sentinel'
  Assert ([IO.File]::ReadAllText($protectedTrial) -eq 'trial sentinel') 'installer lifecycle preserves trial sentinel'
  Assert ((ReadValue $namespace TrialState) -eq 'trial registry sentinel') 'installer lifecycle preserves trial registry sentinel'
  Write-Output "Installer fixture: $count/$count passed; isolated AppId/HKCU namespace, PE files; runtime/running-Excel mocks. Actual clean-machine acceptance remains separate."
} finally {
  [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($namespace,$false)
  [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($arp,$false)
}
