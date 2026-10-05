param(
    [string]$NuGetRoot = (Join-Path $env:USERPROFILE '.nuget/packages'),
    [string]$DebugPublishPath,
    [string]$ReleasePublishPath
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot
$tasks = [Reflection.Assembly]::LoadFrom((Join-Path $NuGetRoot 'exceldna.addin/1.9.0/tools/net6.0-windows7.0/ExcelDna.AddIn.Tasks.dll'))
$resourceType = $tasks.GetType('ResourceHelper')
$flags = [Reflection.BindingFlags]'Public,NonPublic,Static'
function ResourceCall([string]$name, [object[]]$arguments) {
    for ($index = 0; $index -lt $arguments.Length; $index++) { $arguments[$index] = $arguments[$index].PSObject.BaseObject }
    $resourceType.GetMethod($name, $flags).Invoke($null, $arguments)
}
function ReadResource([IntPtr]$module, [string]$name, [string]$type) {
    $resource = ResourceCall 'FindResource' @($module, $name, $type)
    if ($resource -eq [IntPtr]::Zero) { throw "Missing $type/$name" }
    $size = ResourceCall 'SizeofResource' @($module, $resource)
    $data = ResourceCall 'LockResource' @((ResourceCall 'LoadResource' @($module, $resource)))
    $bytes = [byte[]]::new($size)
    [Runtime.InteropServices.Marshal]::Copy($data, $bytes, 0, $size)
    return ,$bytes
}
$appConfig = [IO.File]::ReadAllBytes((Join-Path $repo 'MonteCarlo.Excel/MonteCarlo.Excel/App.config'))
foreach ($configuration in @('Debug','Release')) {
    foreach ($bitness in @('','64')) {
        $publish = if ($configuration -eq 'Debug' -and $DebugPublishPath) { $DebugPublishPath }
            elseif ($configuration -eq 'Release' -and $ReleasePublishPath) { $ReleasePublishPath }
            else { Join-Path $repo "artifacts/diagnostics-$($configuration.ToLower())" }
        $path = Join-Path $publish "MonteCarlo.Excel-AddIn$bitness-packed.xll"
        $module = ResourceCall 'LoadXllResources' @($path)
        try {
            $dnaBytes = ReadResource $module '__MAIN__' 'DNA'
            [xml]$dna = [Text.Encoding]::UTF8.GetString($dnaBytes)
            if ($dna.DnaLibrary.ExternalLibrary.ExplicitExports -ne 'true') { throw 'Implicit exports in packed DNA.' }
            $config = ReadResource $module '__MAIN__' 'CONFIG'
            foreach ($bytes in @($config, [IO.File]::ReadAllBytes($path + '.config'))) {
                if ([Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([byte[]]$bytes)) -ne
                    [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($appConfig))) { throw 'Configuration mismatch.' }
            }
            foreach ($name in @('Core','Excel')) {
                $compressed = ReadResource $module "MONTECARLO.$($name.ToUpper())" 'ASSEMBLY_LZMA'
                $arguments = [object[]]::new(1); $arguments[0] = $compressed
                $bytes = $tasks.GetType('SevenZip.Compression.LZMA.SevenZipHelper').GetMethod('Decompress',$flags).Invoke($null, $arguments)
                $source = if ($name -eq 'Core') { Join-Path $repo "MonteCarlo.Core/bin/$configuration/net10.0/MonteCarlo.Core.dll" }
                    else { Join-Path $repo "MonteCarlo.Excel/MonteCarlo.Excel/bin/$configuration/net10.0-windows/MonteCarlo.Excel.dll" }
                if ([Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([byte[]]$bytes)) -ne (Get-FileHash $source).Hash) { throw 'Stale packed assembly.' }
            }
            $pe = [IO.File]::ReadAllBytes($path)
            $machine = [BitConverter]::ToUInt16($pe, [BitConverter]::ToInt32($pe, 0x3c) + 4)
            if ($machine -ne $(if ($bitness) { 0x8664 } else { 0x14c })) { throw 'Wrong XLL architecture.' }
            if ((ResourceCall 'FindResource' @($module, 'MICROSOFT.OFFICE.INTEROP.EXCEL', 'ASSEMBLY_LZMA')) -ne [IntPtr]::Zero) {
                throw 'External Office PIA dependency still packed instead of embedded.'
            }
            Write-Output "PASS $configuration x$(if($bitness){64}else{86}): ExplicitExports, embedded CONFIG, sidecar hash, final Core/Excel assembly hashes, PE architecture, embedded COM types."
        } finally { [void](ResourceCall 'FreeXllResources' @($module)) }
    }
}
