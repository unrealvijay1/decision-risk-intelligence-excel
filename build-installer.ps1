<#
.SYNOPSIS
Builds and verifies the single customer Setup EXE from fresh Release artifacts.
.DESCRIPTION
Requires .NET 10 SDK and Inno Setup. Runtime downloads are pinned and signature checked.
CodeSigningScript receives -Stage Assemblies|Xll|Installer -Paths <files>.
LiveExcelChecks uses a private Excel process and developer fixture; it is not a clean VM test.
#>
[CmdletBinding()]
param(
  [string]$IsccPath='C:\Program Files\Inno Setup 7\ISCC.exe',
  [string]$RuntimeCache=(Join-Path $PSScriptRoot 'artifacts/runtime-cache'),
  [string]$CodeSigningScript,
  [switch]$RequireSigned,
  [switch]$LiveExcelChecks
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
Set-Location $PSScriptRoot
function Run([string]$Program,[string[]]$Arguments) {
  & $Program @Arguments
  if($LASTEXITCODE -ne 0){throw "$Program failed with exit code $LASTEXITCODE"}
}
function Sign([string]$Stage,[string[]]$Paths) {
  if($CodeSigningScript) {
    & $CodeSigningScript -Stage $Stage -Paths $Paths
    if(-not $?){throw "Signing hook failed: $Stage"}
    foreach($path in $Paths){if((Get-AuthenticodeSignature $path).Status -ne 'Valid'){throw "Signing verification failed: $path"}}
  }
}
if($RequireSigned -and -not $CodeSigningScript){throw 'RequireSigned requires a real code-signing hook.'}
if($CodeSigningScript -and -not (Test-Path -LiteralPath $CodeSigningScript -PathType Leaf)){throw 'Signing hook not found.'}
if(-not (Test-Path -LiteralPath $IsccPath -PathType Leaf)){throw 'Install Inno Setup or supply -IsccPath.'}
[xml]$props=Get-Content Directory.Build.props
$version=[string]$props.Project.PropertyGroup[0].ProductVersion
if($version -notmatch '^\d+\.\d+\.\d+$'){throw 'Invalid authoritative ProductVersion'}
$root=Join-Path $PSScriptRoot ('artifacts/installer-'+[guid]::NewGuid().ToString('N'))
$publish=Join-Path $root 'publish'
$stage=Join-Path $root 'payload'
$dist=Join-Path $PSScriptRoot 'dist'
[void](New-Item -ItemType Directory -Force $root,$stage,$dist,$RuntimeCache)
Start-Transcript -Path (Join-Path $root 'build.log') | Out-Null
try {
$buildProperty='-p:InstallerBuildRoot='+$root.Replace('\','/')
$publishProperty='-p:ExcelDnaPublishPath='+$publish.Replace('\','/')
Run dotnet @('clean','MonteCarlo.Excel/MonteCarlo.Excel.slnx','-c','Release',$buildProperty,$publishProperty)
Run dotnet @('build','MonteCarlo.Excel/MonteCarlo.Excel.slnx','-c','Release',$buildProperty,$publishProperty)
Run dotnet @('test','MonteCarlo.Core.Tests/MonteCarlo.Core.Tests.csproj','-c','Release','--no-build','--no-restore',$buildProperty,'--logger','trx','--results-directory',(Join-Path $root 'test-results'))
Run dotnet @('build','Installer/Checks/Checks.csproj','-c','Release')
if($CodeSigningScript) {
  $assemblies=@((Join-Path $root 'MonteCarlo.Excel/bin/Release/net10.0-windows/MonteCarlo.Excel.dll'),(Join-Path $root 'MonteCarlo.Core/bin/Release/net10.0/MonteCarlo.Core.dll'))
  Sign Assemblies $assemblies
  Copy-Item -LiteralPath $assemblies[1] -Destination (Join-Path $root 'MonteCarlo.Excel/bin/Release/net10.0-windows/MonteCarlo.Core.dll')
  # The pack target consumes the already-signed outputs; do not compile again.
  Run dotnet @('msbuild','MonteCarlo.Excel/MonteCarlo.Excel/MonteCarlo.Excel.csproj','-t:ExcelDnaPack','-p:Configuration=Release',$buildProperty,$publishProperty)
}
$xlls=@((Join-Path $publish 'MonteCarlo.Excel-AddIn-packed.xll'),(Join-Path $publish 'MonteCarlo.Excel-AddIn64-packed.xll'))
Sign Xll $xlls
& ./Installer/Verify-Payload.ps1 -PublishPath $publish -BuildRoot $root -Version $version
if(-not $?){throw 'Payload verification failed'}
foreach($xll in $xlls){Copy-Item -LiteralPath $xll -Destination $stage; Copy-Item -LiteralPath ($xll+'.config') -Destination $stage}
Copy-Item -LiteralPath Installer/CustomerGuide.txt -Destination $stage
$runtime=Get-Content Installer/runtime-manifest.json -Raw | ConvertFrom-Json
if($runtime.version -ne '10.0.12'){throw 'Update the installer runtime minimum together with runtime-manifest.json'}
foreach($file in $runtime.files) {
  $cache=Join-Path $RuntimeCache $file.name
  if(-not (Test-Path -LiteralPath $cache)) {Invoke-WebRequest -Uri $file.url -OutFile $cache}
  if((Get-FileHash $cache -Algorithm SHA512).Hash -ne $file.sha512){throw "Runtime checksum mismatch: $cache"}
  $signature=Get-AuthenticodeSignature $cache
  if($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation'){throw "Runtime Microsoft signature invalid: $cache"}
  Copy-Item -LiteralPath $cache -Destination $stage
}
$allow=@('MonteCarlo.Excel-AddIn-packed.xll','MonteCarlo.Excel-AddIn-packed.xll.config','MonteCarlo.Excel-AddIn64-packed.xll','MonteCarlo.Excel-AddIn64-packed.xll.config','CustomerGuide.txt')+@($runtime.files.name)
$actual=@(Get-ChildItem -LiteralPath $stage -File | Select-Object -ExpandProperty Name)
if(@(Compare-Object ($allow|Sort-Object) ($actual|Sort-Object)).Count -ne 0){throw 'Customer payload allowlist mismatch'}
& ./Installer/Test-RuntimeDetection.ps1 -PayloadDir $stage -IsccPath $IsccPath -Version $version
if(-not $?){throw 'Real runtime detection verification failed'}
if($LiveExcelChecks) {
  Run dotnet @('build','MonteCarlo.Excel.StartupChecks/MonteCarlo.Excel.StartupChecks.csproj','-c','Release',('-p:ExcelDnaPublishPath='+(Join-Path $PSScriptRoot 'artifacts/diagnostics-fixture')))
}
& ./Installer/Test-Installer.ps1 -PayloadDir $stage -IsccPath $IsccPath -Version $version -LiveStartup:$LiveExcelChecks
if(-not $?){throw 'Installer lifecycle verification failed'}
Run $IsccPath @(('/DProductVersion='+$version),('/DPayloadDir='+$stage),('/DOutputDirectory='+(Join-Path $root 'output')),'Installer/MonteCarloForExcel.iss')
$exe=Join-Path $root "output/MonteCarloForExcel-Setup-$version.exe"
Sign Installer @($exe)
$info=Get-Item -LiteralPath $exe
if($info.Length -lt 1000000 -or $info.VersionInfo.ProductVersion.Trim() -ne $version -or $info.VersionInfo.FileVersion.Trim() -ne ($version+'.0')){throw 'Setup EXE version/size verification failed'}
$destination=Join-Path $dist "MonteCarloForExcel-Setup-$version.exe"
Copy-Item -LiteralPath $exe -Destination $destination
$manifest=[ordered]@{
  version=$version; installer=$destination; bytes=$info.Length; sha256=(Get-FileHash $exe).Hash;
  signature=(Get-AuthenticodeSignature $exe).Status.ToString(); buildRoot=$root;
  payload=@(Get-ChildItem -LiteralPath $stage -File | ForEach-Object { [ordered]@{name=$_.Name;bytes=$_.Length;sha256=(Get-FileHash $_.FullName).Hash} });
  coreTests='Complete Release suite passed; see TRX in buildRoot'; activationChecks='Packed assembly contains no activation types, resources or dependencies';
  installerTests='Isolated installer lifecycle fixture passed; process/runtime mocks, actual PE architecture and payload checks';
  installedCopyLiveExcelChecked=[bool]$LiveExcelChecks;
  cleanMachineAcceptance='Pending: disposable Windows + x86/x64 Excel, no activation state, upgrade/reinstall/uninstall and prerequisite UAC';
}
$manifest | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $dist "MonteCarloForExcel-$version-manifest.json")
Write-Output "READY FOR ACCEPTANCE: $destination ($($info.Length) bytes), version $version, signature $($manifest.signature)."
} finally { Stop-Transcript | Out-Null }
