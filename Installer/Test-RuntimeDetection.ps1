param([Parameter(Mandatory)][string]$PayloadDir,[string]$IsccPath='C:\Program Files\Inno Setup 7\ISCC.exe',[string]$Version='0.2.1')
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot
$root=Join-Path $repo ('artifacts/runtime-detection-'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Force $root)
$source=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'MonteCarloForExcel.iss')).Replace("`r`n","`n")
$source=$source.Replace('procedure InitializeWizard();',"function RuntimeInstalled(): Boolean; forward;`nprocedure InitializeWizard();")
$anchor="function ShouldSkipPage(PageID: Integer): Boolean;"
# Inject a read-only probe before wizard initialization returns; the original
# RuntimeInstalled implementation runs unchanged against real HKLM registry views.
$source=$source.Replace("end;`n"+$anchor,"  Log('RUNTIME_PROBE architecture='+ExcelArchitecture+'; installed='+IntToStr(Ord(RuntimeInstalled())));`n  Abort;`nend;`n"+$anchor)
if($source -notmatch 'RUNTIME_PROBE'){throw 'Probe injection failed'}
$source=$source.Replace('Compression=lzma2','Compression=none').Replace('SolidCompression=yes','SolidCompression=no')
$source=$source.Replace('Major := MS shr 16;','Major := 16;')
$script=Join-Path $root 'Probe.iss'; [IO.File]::WriteAllText($script,$source)
& $IsccPath ('/DProductVersion='+$Version) ('/DPayloadDir='+$PayloadDir) ('/DOutputDirectory='+$root) $script | Set-Content (Join-Path $root 'compile.log')
if($LASTEXITCODE -ne 0){throw 'Probe compile failed'}
foreach($arch in @('x86','x64')) {
  $registry=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::Registry32)
  $key=$registry.OpenSubKey('SOFTWARE\dotnet\Setup\InstalledVersions\'+$arch+'\sharedfx\Microsoft.WindowsDesktop.App')
  $expected=0
  try { if($key){foreach($name in $key.GetValueNames()){if($name -match '^10\.0\.(\d+)$' -and [int]$Matches[1] -ge 12){$expected=1}}} }
  finally {if($key){$key.Dispose()};$registry.Dispose()}
  $directory=Join-Path $root $arch; [void](New-Item -ItemType Directory -Force $directory)
  $excel=Join-Path $directory 'EXCEL.EXE'
  $bytes=[IO.File]::ReadAllBytes((Join-Path $repo 'Installer/Checks/bin/Release/net10.0-windows/Checks.exe'))
  $offset=[BitConverter]::ToInt32($bytes,0x3c)
  if($arch -eq 'x86'){$bytes[$offset+4]=0x4c;$bytes[$offset+5]=0x01}else{$bytes[$offset+4]=0x64;$bytes[$offset+5]=0x86}
  [IO.File]::WriteAllBytes($excel,$bytes)
  $log=Join-Path $directory 'probe.log'
  $process=Start-Process -FilePath (Join-Path $root "MonteCarloForExcel-Setup-$Version.exe") -ArgumentList @('/VERYSILENT','/SUPPRESSMSGBOXES',('/EXCELPATH="'+$excel+'"'),('/LOG="'+$log+'"')) -WindowStyle Hidden -PassThru
  if(-not $process.WaitForExit(30000)){throw 'Read-only runtime probe timed out'}
  if([IO.File]::ReadAllText($log) -notmatch ("RUNTIME_PROBE architecture="+$arch+"; installed="+$expected)){throw "$arch real runtime registration mismatch"}
  Write-Output "PASS original RuntimeInstalled matches real $arch Desktop Runtime registration (installed=$expected); probe aborts before installation/registration."
}
Write-Output $root
