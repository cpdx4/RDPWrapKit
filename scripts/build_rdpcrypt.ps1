<#
.SYNOPSIS
    Builds RdpSignTool.exe ahead-of-time for the installer.

.DESCRIPTION
    Compiles RdpSignTool.exe — a standalone C# console EXE that signs .rdp
    files without ANY PowerShell overhead.  Called directly from Inno Setup
    via Exec(), eliminating ~5 seconds of PS startup + -EncodedCommand +
    Add-Type overhead.

    Uses the .NET Framework csc.exe (C# compiler) that ships with the .NET
    Framework runtime or Windows SDK.

    Output:  ..\output\RdpSignTool.exe

.EXAMPLE
    .\build_rdpcrypt.ps1
#>

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $PSCommandPath
$ProjectRoot = Split-Path -Parent $ScriptDir
$OutputDir  = Join-Path $ProjectRoot 'output'

# Ensure output directory exists
if (-not (Test-Path $OutputDir)) {
    $null = New-Item -ItemType Directory -Path $OutputDir -Force
}

# Locate csc.exe — check common .NET Framework SDK / runtime paths
$cscPaths = @(
    "${env:SystemRoot}\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "${env:SystemRoot}\Microsoft.NET\Framework\v4.0.30319\csc.exe",
    "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\csc.exe",
    "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.7.2 Tools\csc.exe",
    "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.7 Tools\csc.exe",
    "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.6.2 Tools\csc.exe",
    "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.6.1 Tools\csc.exe",
    "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.6 Tools\csc.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\Roslyn\csc.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Community\MSBuild\15.0\Bin\Roslyn\csc.exe"
)

$CscExe = $null
foreach ($candidate in $cscPaths) {
    if (Test-Path $candidate) {
        $CscExe = $candidate
        break
    }
}

if (-not $CscExe) {
    try {
        $CscExe = (Get-Command 'csc.exe' -ErrorAction Stop).Source
    } catch {
        throw "Could not locate csc.exe. Install .NET Framework SDK or Visual Studio Build Tools."
    }
}

Write-Host "Using C# compiler: $CscExe"
Write-Host "Output directory: $OutputDir"

$references = @(
    '/reference:System.dll',
    '/reference:System.Security.dll',
    '/reference:System.Core.dll'
)

# ---------------------------------------------------------------
# Build: RdpSignTool.exe  (standalone signer, no PS needed)
# ---------------------------------------------------------------
$exeSource = Join-Path $ScriptDir 'RdpSignTool.cs'
$exeOutput = Join-Path $OutputDir 'RdpSignTool.exe'

Write-Host ""
Write-Host "=== Building RdpSignTool.exe ==="
Write-Host "Source: $exeSource"
Write-Host "Output: $exeOutput"
Write-Host "Version: 1.0.0.0"

$exeArgs = @(
    '/target:exe', '/nologo', '/debug-', '/optimize+'
) + $references + @(
    "/out:$exeOutput",
    $exeSource
)

Write-Host "Compiling... " -NoNewline
$proc = Start-Process -FilePath $CscExe -ArgumentList $exeArgs -NoNewWindow -Wait -PassThru
if ($proc.ExitCode -eq 0) {
    $size = (Get-Item $exeOutput).Length
    Write-Host "OK ($size bytes)"
    Write-Host ""
    Write-Host "Build completed successfully."
} else {
    Write-Host "FAILED (exit: $($proc.ExitCode))"
    exit $proc.ExitCode
}
