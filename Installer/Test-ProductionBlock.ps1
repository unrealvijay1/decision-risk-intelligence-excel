param([Parameter(Mandatory)][string]$InstallerPath,[Parameter(Mandatory)][string]$ExcelPath)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot
$root=Join-Path $repo ('artifacts/installer-block-'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Force $root)
# This probe must only run while Excel already exists. It never closes those processes.
$existing=@(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" | Select-Object -ExpandProperty ProcessId)
if($existing.Count -eq 0){throw 'Running-Excel probe requires an already running Excel. No Setup launched.'}
function Snapshot {
  $result=[ordered]@{}
  foreach($path in @('Software\Microsoft\Office\16.0\Excel\Options','Software\MonteCarloExcel\Installer','Software\Microsoft\Windows\CurrentVersion\Uninstall\{9B4D8E57-7B33-4C68-9D1D-91D5E4D8F001}_is1')) {
    $key=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($path)
    try {
      $data=[ordered]@{}
      if($key){foreach($name in ($key.GetValueNames()|Sort-Object)){$data[$name]=$key.GetValue($name)}}
      $result[$path]=[ordered]@{exists=($null -ne $key);values=$data}
    } finally {if($key){$key.Dispose()}}
  }
  return ($result | ConvertTo-Json -Depth 8 -Compress)
}
$before=Snapshot
[IO.File]::WriteAllText((Join-Path $root 'registry-before.json'),$before)
$log=Join-Path $root 'blocked.log'
$target=Join-Path $root 'must-not-install'
$p=Start-Process -FilePath $InstallerPath -ArgumentList @('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART',('/EXCELPATH="'+$ExcelPath+'"'),('/DIR="'+$target+'"'),('/LOG="'+$log+'"')) -WindowStyle Hidden -PassThru
if(-not $p.WaitForExit(30000)){throw 'Production gate probe timed out; no customer process will be stopped.'}
$after=Snapshot
[IO.File]::WriteAllText((Join-Path $root 'registry-after.json'),$after)
if($p.ExitCode -eq 0 -or $before -ne $after -or (Test-Path -LiteralPath (Join-Path $target 'MonteCarloForExcel.xll'))){throw 'Running-Excel block failed: inspect snapshots/log before further testing.'}
if([IO.File]::ReadAllText($log) -notmatch 'Close all Microsoft Excel instances'){throw 'Expected running-Excel message missing.'}
$report="PASS production Setup running-Excel gate; exit $($p.ExitCode); registry before/after identical; no XLL installed; no Excel process operated on."
[IO.File]::WriteAllText((Join-Path $root 'report.txt'),$report)
Write-Output $report
Write-Output $root
