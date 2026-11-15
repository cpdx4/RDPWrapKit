@echo off
setlocal EnableExtensions
REM ===========================================================================
REM RDPWrapKit Build Wrapper
REM ===========================================================================
REM
REM 1. Compiles RdpSignTool.exe from scripts\RdpSignTool.cs (C# source)
REM 2. Compiles RDPWrapKit-Setup.exe from RDPWrapKit.iss (Inno Setup)
REM
REM Step 1 invokes csc.exe directly. It used to shell out to
REM "powershell.exe -ExecutionPolicy Bypass -File scripts\build_rdpcrypt.ps1",
REM which put an execution-policy bypass on a command line in every build (and
REM in CI logs). The compiler lookup order now mirrors the ISCC lookup below.
REM The .ps1 is kept for developers who prefer it, but the build no longer
REM depends on PowerShell at all.
REM
REM Usage:
REM   build.bat              full build (1-step)
REM   build.bat signonly     only build RdpSignTool.exe
REM ===========================================================================

set ROOT=%~dp0
set SIGNONLY=0
if /i "%~1"=="signonly" set SIGNONLY=1

echo.
echo === Step 1: Build RdpSignTool.exe ===
if not exist "%ROOT%scripts\RdpSignTool.cs" (
    echo ERROR: scripts\RdpSignTool.cs not found in %ROOT%scripts
    exit /b 1
)
if not exist "%ROOT%output" mkdir "%ROOT%output"

set CSC_EXE=
if exist "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" set CSC_EXE=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not defined CSC_EXE if exist "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe" set CSC_EXE=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe
if not defined CSC_EXE if exist "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\csc.exe" set CSC_EXE=C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\csc.exe
if not defined CSC_EXE if exist "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.7.2 Tools\csc.exe" set CSC_EXE=C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.7.2 Tools\csc.exe
if not defined CSC_EXE if exist "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.6.2 Tools\csc.exe" set CSC_EXE=C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.6.2 Tools\csc.exe
if not defined CSC_EXE for %%G in (csc.exe) do if not "%%~$PATH:G"=="" set CSC_EXE=%%~$PATH:G

if not defined CSC_EXE (
    echo ERROR: csc.exe not found.
    echo Looked in the .NET Framework 4.x runtime folder, the Windows SDK NETFX
    echo tools and PATH. Install .NET Framework 4.x or Visual Studio Build Tools.
    exit /b 1
)

echo Using: %CSC_EXE%
"%CSC_EXE%" /nologo /target:exe /optimize+ ^
    /reference:System.dll /reference:System.Security.dll /reference:System.Core.dll ^
    /out:"%ROOT%output\RdpSignTool.exe" "%ROOT%scripts\RdpSignTool.cs"
if errorlevel 1 (
    echo ERROR: RdpSignTool.exe build failed
    exit /b 1
)

if %SIGNONLY% equ 1 (
    echo.
    echo RdpSignTool.exe built successfully.
    exit /b 0
)

echo.
echo === Step 2: Compile Installer ===

REM Try to find ISCC.exe in common locations (prefer Inno Setup 7, fall back to 6)
set ISCC_EXE=
if exist "C:\Program Files\Inno Setup 7\ISCC.exe" set ISCC_EXE=C:\Program Files\Inno Setup 7\ISCC.exe
if not defined ISCC_EXE if exist "C:\Program Files (x86)\Inno Setup 7\ISCC.exe" set ISCC_EXE=C:\Program Files (x86)\Inno Setup 7\ISCC.exe
if not defined ISCC_EXE if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set ISCC_EXE=C:\Program Files (x86)\Inno Setup 6\ISCC.exe
if not defined ISCC_EXE if exist "C:\Program Files\Inno Setup 6\ISCC.exe" set ISCC_EXE=C:\Program Files\Inno Setup 6\ISCC.exe
if not defined ISCC_EXE where ISCC.exe >nul 2>nul && set ISCC_EXE=ISCC.exe

if not defined ISCC_EXE (
    echo ERROR: Inno Setup compiler not found.
    echo Looked for Inno Setup 7 then Inno Setup 6.
    echo Install from https://jrsoftware.org/isdl.php
    exit /b 1
)

echo Using: %ISCC_EXE%
"%ISCC_EXE%" "%ROOT%RDPWrapKit.iss"
if errorlevel 1 (
    echo ERROR: Installer compilation failed
    exit /b 1
)

echo.
echo ========================================
echo  Build completed successfully!
echo ========================================
