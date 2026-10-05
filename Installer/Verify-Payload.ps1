param([Parameter(Mandatory)][string]$PublishPath,[Parameter(Mandatory)][string]$BuildRoot,[Parameter(Mandatory)][string]$Version,[string]$Configuration='Release')
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot
$tasks=[Reflection.Assembly]::LoadFrom((Join-Path $env:USERPROFILE '.nuget/packages/exceldna.addin/1.9.0/tools/net6.0-windows7.0/ExcelDna.AddIn.Tasks.dll'))
$resourceType=$tasks.GetType('ResourceHelper')
$flags=[Reflection.BindingFlags]'Public,NonPublic,Static'
if (-not ('CustomerResources' -as [type])) {
Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class CustomerResources {
 delegate bool Callback(IntPtr module, IntPtr type, IntPtr name, IntPtr param);
 [DllImport("kernel32.dll", CharSet=CharSet.Unicode, EntryPoint="EnumResourceNamesW", SetLastError=true)]
 static extern bool EnumNames(IntPtr module, string type, Callback callback, IntPtr param);
 public static string[] Names(IntPtr module, string type) {
  var names=new List<string>();
  EnumNames(module,type,(m,t,n,p)=>{names.Add(n.ToInt64() <= 65535 ? "#"+n.ToInt64() : Marshal.PtrToStringUni(n)!);return true;},IntPtr.Zero);
  return names.ToArray();
 }
}
'@
}
function ResourceCall([string]$name,[object[]]$arguments) {
  for($i=0;$i -lt $arguments.Length;$i++){ $arguments[$i]=$arguments[$i].PSObject.BaseObject }
  $resourceType.GetMethod($name,$flags).Invoke($null,$arguments)
}
function ReadResource([IntPtr]$module,[string]$name,[string]$type) {
  $res=ResourceCall 'FindResource' @($module,$name,$type)
  if($res -eq [IntPtr]::Zero){throw "Missing packed resource $type/$name"}
  $size=ResourceCall 'SizeofResource' @($module,$res)
  $ptr=ResourceCall 'LockResource' @((ResourceCall 'LoadResource' @($module,$res)))
  $bytes=[byte[]]::new($size); [Runtime.InteropServices.Marshal]::Copy($ptr,$bytes,0,$size)
  return ,$bytes
}
$expectedConfig=Get-FileHash (Join-Path $repo 'MonteCarlo.Excel/MonteCarlo.Excel/App.config')
$inventory=@()
foreach($suffix in @('','64')) {
  $path=Join-Path $PublishPath "MonteCarlo.Excel-AddIn$suffix-packed.xll"
  $versionInfo=(Get-Item -LiteralPath $path).VersionInfo
  if($versionInfo.FileVersion.Trim() -ne ($Version+'.0') -or $versionInfo.ProductVersion.Trim() -ne $Version){throw 'Packed XLL product/file version mismatch'}
  $pe=[IO.File]::ReadAllBytes($path)
  if([BitConverter]::ToUInt16($pe,[BitConverter]::ToInt32($pe,0x3c)+4) -ne $(if($suffix){0x8664}else{0x14c})){throw 'XLL architecture mismatch'}
  $module=ResourceCall 'LoadXllResources' @($path)
  try {
    $assemblyNames=@([CustomerResources]::Names($module,'ASSEMBLY_LZMA'))
    $allowedAssemblies=@('MONTECARLO.CORE','MONTECARLO.EXCEL','EXCELDNA.INTEGRATION','EXCELDNA.LOADER','EXCELDNA.MANAGEDHOST')
    foreach($packedName in $assemblyNames) {
      if($packedName -notin $allowedAssemblies){throw "Unapproved packed assembly: $packedName"}
    }
    Write-Output ('Packed managed inventory: '+($assemblyNames -join ', '))
    [xml]$dna=[Text.Encoding]::UTF8.GetString((ReadResource $module '__MAIN__' 'DNA'))
    if($dna.DnaLibrary.ExternalLibrary.ExplicitExports -ne 'true'){throw 'Implicit exports enabled'}
    $config=ReadResource $module '__MAIN__' 'CONFIG'
    if([Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($config)) -ne $expectedConfig.Hash -or (Get-FileHash ($path+'.config')).Hash -ne $expectedConfig.Hash){throw 'Packed/sidecar configuration mismatch'}
    if(([xml][Text.Encoding]::UTF8.GetString($config)).configuration.monteCarloDiagnostics.verbose -ne 'false'){throw 'Developer diagnostics enabled'}
    foreach($name in @('Core','Excel')) {
      $compressed=ReadResource $module "MONTECARLO.$($name.ToUpper())" 'ASSEMBLY_LZMA'
      $arguments=[object[]]::new(1); $arguments[0]=$compressed
      [byte[]]$bytes=$tasks.GetType('SevenZip.Compression.LZMA.SevenZipHelper').GetMethod('Decompress',$flags).Invoke($null,$arguments)
      $framework=if($name -eq 'Core'){'net10.0'}else{'net10.0-windows'}
      $source=Join-Path $BuildRoot "MonteCarlo.$name/bin/$Configuration/$framework/MonteCarlo.$name.dll"
      if([Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes)) -ne (Get-FileHash $source).Hash){throw 'Packed assembly is stale'}
      $text=[Text.Encoding]::UTF8.GetString($bytes)
      if($text.Contains('-----BEGIN PRIVATE KEY-----') -or $text.Contains('-----BEGIN RSA PRIVATE KEY-----')){throw 'Private PEM in packed assembly'}
      $extracted=Join-Path $BuildRoot "verified-$suffix/MonteCarlo.$name.dll"
      [void](New-Item -ItemType Directory -Force (Split-Path $extracted)); [IO.File]::WriteAllBytes($extracted,$bytes)
      if($name -eq 'Excel') {
        & (Join-Path $repo 'Installer/Checks/bin/Release/net10.0-windows/Checks.exe') $extracted $repo $Version
        if($LASTEXITCODE -ne 0){throw 'Activation-free runtime verification failed'}
      }
    }
    foreach($forbidden in @('MONTECARLO.LICENSEGENERATOR','MICROSOFT.OFFICE.INTEROP.EXCEL','OFFICE')) {
      if((ResourceCall 'FindResource' @($module,$forbidden,'ASSEMBLY_LZMA')) -ne [IntPtr]::Zero){throw "Forbidden packed dependency $forbidden"}
    }
    $inventory += [pscustomobject]@{Architecture=$(if($suffix){'x64'}else{'x86'});Path=$path;SHA256=(Get-FileHash $path).Hash;Bytes=(Get-Item $path).Length}
    Write-Output "PASS packed $(if($suffix){'x64'}else{'x86'}): PE, explicit exports, production configuration, fresh assemblies, no activation types/resources."
  } finally {[void](ResourceCall 'FreeXllResources' @($module))}
}
$inventory | ConvertTo-Json | Set-Content (Join-Path $BuildRoot 'packed-verification.json')
