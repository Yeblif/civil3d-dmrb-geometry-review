$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "DMRB Superelevation Exporter - No-SDK NETLOAD Build" -ForegroundColor Cyan
Write-Host "----------------------------------------------------"
Write-Host "Uses the Windows C# compiler plus the .NET 8 runtime already on the PC."
Write-Host "No SDK is installed and nothing is written to Autodesk ApplicationPlugins."
Write-Host ""

$source = Join-Path $PSScriptRoot "src\Plugin.cs"
if (-not (Test-Path $source)) {
    throw "Source file not found: $source"
}

# Locate Civil 3D / AutoCAD managed API assemblies.
$acadRoot = "C:\Program Files\Autodesk\AutoCAD 2026"
$acMgd = Join-Path $acadRoot "AcMgd.dll"
$acDbMgd = Join-Path $acadRoot "AcDbMgd.dll"
$acCoreMgd = Join-Path $acadRoot "AcCoreMgd.dll"
$aeccDbMgd = Join-Path $acadRoot "C3D\AeccDbMgd.dll"

if (-not (Test-Path $acMgd) -or
    -not (Test-Path $acDbMgd) -or
    -not (Test-Path $acCoreMgd) -or
    -not (Test-Path $aeccDbMgd)) {

    Write-Host "Standard Civil 3D 2026 paths were not complete; searching Autodesk Program Files..." -ForegroundColor Yellow

    $autodeskRoot = "C:\Program Files\Autodesk"
    if (-not (Test-Path $autodeskRoot)) {
        throw "C:\Program Files\Autodesk was not found."
    }

    function Find-One([string]$name) {
        $item = Get-ChildItem -Path $autodeskRoot -Filter $name -File -Recurse -ErrorAction SilentlyContinue |
                Select-Object -First 1
        if ($null -eq $item) {
            throw "Could not find $name under $autodeskRoot"
        }
        return $item.FullName
    }

    $acMgd = Find-One "AcMgd.dll"
    $acDbMgd = Find-One "AcDbMgd.dll"
    $acCoreMgd = Find-One "AcCoreMgd.dll"

    $aeccCandidates = Get-ChildItem -Path $autodeskRoot -Filter "AeccDbMgd.dll" -File -Recurse -ErrorAction SilentlyContinue |
                      Where-Object { $_.FullName -match "AutoCAD 2026" } |
                      Select-Object -First 1

    if ($null -eq $aeccCandidates) {
        throw "Could not find the Civil 3D 2026 AeccDbMgd.dll."
    }
    $aeccDbMgd = $aeccCandidates.FullName
}

# Civil 3D Alignment derives through the AEC managed API, so AecBaseMgd is
# also required at compile time.
$aecBaseMgd = $null
$aecLikely = @(
    (Join-Path $acadRoot "AecBaseMgd.dll"),
    (Join-Path $acadRoot "AEC\AecBaseMgd.dll"),
    (Join-Path $acadRoot "C3D\AecBaseMgd.dll")
)

foreach ($candidate in $aecLikely) {
    if (Test-Path $candidate) {
        $aecBaseMgd = $candidate
        break
    }
}

if ($null -eq $aecBaseMgd) {
    $aecCandidate = Get-ChildItem -Path $acadRoot -Filter "AecBaseMgd.dll" -File -Recurse -ErrorAction SilentlyContinue |
                    Select-Object -First 1
    if ($null -ne $aecCandidate) {
        $aecBaseMgd = $aecCandidate.FullName
    }
}

if ($null -eq $aecBaseMgd) {
    throw "Could not find AecBaseMgd.dll under the Civil 3D / AutoCAD 2026 installation."
}

Write-Host "AutoCAD API: $acMgd"
Write-Host "AEC API: $aecBaseMgd"
Write-Host "Civil 3D API: $aeccDbMgd"

# Locate .NET 8 shared runtime assemblies.
$dotnetShared = Join-Path $env:ProgramFiles "dotnet\shared"
$coreBase = Join-Path $dotnetShared "Microsoft.NETCore.App"
$desktopBase = Join-Path $dotnetShared "Microsoft.WindowsDesktop.App"

function Latest-V8Folder([string]$base) {
    if (-not (Test-Path $base)) {
        return $null
    }

    $candidates = Get-ChildItem -Path $base -Directory -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -match '^8\.' } |
                  Sort-Object { [version]$_.Name } -Descending

    if ($candidates.Count -eq 0) {
        return $null
    }

    return $candidates[0].FullName
}

$coreDir = Latest-V8Folder $coreBase
$desktopDir = Latest-V8Folder $desktopBase

# Civil 3D normally uses the globally installed .NET 8 Desktop runtime.
# Some managed workstations keep the runtime alongside Autodesk instead, so
# fall back to the AutoCAD installation if the global folders are unavailable.
if ($null -eq $coreDir) {
    $coreCandidate = Get-ChildItem -Path $acadRoot -Filter "System.Private.CoreLib.dll" -File -Recurse -ErrorAction SilentlyContinue |
                     Select-Object -First 1
    if ($null -ne $coreCandidate) {
        $coreDir = $coreCandidate.Directory.FullName
    }
}

if ($null -eq $desktopDir) {
    $desktopCandidate = Get-ChildItem -Path $acadRoot -Filter "System.Windows.Forms.dll" -File -Recurse -ErrorAction SilentlyContinue |
                        Select-Object -First 1
    if ($null -ne $desktopCandidate) {
        $desktopDir = $desktopCandidate.Directory.FullName
    }
}

if ($null -eq $coreDir) {
    throw "Could not locate .NET 8 core runtime assemblies."
}

if ($null -eq $desktopDir) {
    throw "Could not locate the .NET 8 Windows Desktop runtime assemblies required for the alignment-selection window."
}

Write-Host ".NET Core references: $coreDir"
Write-Host ".NET Desktop references: $desktopDir"

# Locate a compiler. Prefer any Visual Studio/Roslyn csc.exe, otherwise use the
# compiler included with Windows .NET Framework. The source intentionally uses
# conservative C# syntax so this fallback can compile it.
$cscCandidates = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)

$csc = $null
foreach ($candidate in $cscCandidates) {
    if (Test-Path $candidate) {
        $csc = $candidate
        break
    }
}

if ($null -eq $csc) {
    throw "Windows C# compiler csc.exe was not found. No files have been installed or changed."
}

Write-Host "Compiler: $csc"

$outDir = Join-Path $PSScriptRoot "NETLOAD"
if (-not (Test-Path $outDir)) {
    New-Item -Path $outDir -ItemType Directory | Out-Null
}

$outDll = Join-Path $outDir "DMRBSuperelevationExporter.dll"
$outPdb = Join-Path $outDir "DMRBSuperelevationExporter.pdb"
$rsp = Join-Path $PSScriptRoot "build_no_sdk.rsp"

Remove-Item $outDll -Force -ErrorAction SilentlyContinue
Remove-Item $outPdb -Force -ErrorAction SilentlyContinue
Remove-Item $rsp -Force -ErrorAction SilentlyContinue

# Build a de-duplicated reference set by DLL filename.
# Runtime folders also contain native DLLs such as coreclr.dll; csc cannot use
# those as metadata references, so only include files that .NET recognises as
# managed assemblies.
$refs = @{}

function Add-ManagedReferences([string]$folder) {
    Get-ChildItem -Path $folder -Filter "*.dll" -File | ForEach-Object {
        try {
            [void][System.Reflection.AssemblyName]::GetAssemblyName($_.FullName)
            $key = $_.Name.ToLowerInvariant()
            if (-not $refs.ContainsKey($key)) {
                $refs[$key] = $_.FullName
            }
        }
        catch {
            # Native/non-managed DLL: deliberately ignored.
        }
    }
}

Add-ManagedReferences $coreDir

# Desktop contains the real WPF/WinForms assemblies. Override same-named Core
# facades/stubs such as WindowsBase.dll.
Get-ChildItem -Path $desktopDir -Filter "*.dll" -File | ForEach-Object {
    try {
        [void][System.Reflection.AssemblyName]::GetAssemblyName($_.FullName)
        $refs[$_.Name.ToLowerInvariant()] = $_.FullName
    }
    catch {
        # Native/non-managed DLL: deliberately ignored.
    }
}

foreach ($p in @($acCoreMgd, $acDbMgd, $acMgd, $aecBaseMgd, $aeccDbMgd)) {
    $refs[(Split-Path $p -Leaf).ToLowerInvariant()] = $p
}

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("/nostdlib+")
$lines.Add("/target:library")
$lines.Add("/platform:x64")
$lines.Add("/optimize+")
$lines.Add("/debug:pdbonly")
$lines.Add("/utf8output")
$lines.Add('/out:"' + $outDll + '"')

foreach ($path in ($refs.Values | Sort-Object)) {
    $lines.Add('/reference:"' + $path + '"')
}

$lines.Add('"' + $source + '"')
$lines | Set-Content -Path $rsp -Encoding Unicode

Write-Host ""
Write-Host "Compiling..." -ForegroundColor Cyan

& $csc "/noconfig" "@$rsp"
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0 -or -not (Test-Path $outDll)) {
    Write-Host ""
    Write-Host "Compilation failed. Nothing was installed." -ForegroundColor Red
    throw "Compiler returned exit code $exitCode. Copy the errors above into a GitHub issue."
}

Write-Host ""
Write-Host "Build successful." -ForegroundColor Green
Write-Host ""
Write-Host "NETLOAD this file in Civil 3D:"
Write-Host "  $outDll" -ForegroundColor Cyan
Write-Host ""
Write-Host "Then run:"
Write-Host "  DMRBEXPORTSUPER" -ForegroundColor Cyan
Write-Host ""
