@echo off
setlocal EnableExtensions
REM ===========================================================================
REM RDPWrapKit Build Wrapper
REM ===========================================================================
REM
REM 1. Compiles RdpWrapKitHelper.exe from scripts\RdpWrapKitHelper.cs (C# source)
REM 2. Compiles RDPWrapKit-Setup.exe from RDPWrapKit.iss (Inno Setup)
REM
REM Step 1 invokes csc.exe directly with the same compiler lookup order used by
REM the ISCC lookup below, so the build needs no script host.
REM
REM Usage:
REM   build.bat              full build (1-step)
REM   build.bat helperonly   only build RdpWrapKitHelper.exe
REM   build.bat signonly     alias of helperonly (kept for existing scripts)
REM ===========================================================================

set ROOT=%~dp0
set SIGNONLY=0
if /i "%~1"=="helperonly" set SIGNONLY=1
if /i "%~1"=="signonly" set SIGNONLY=1

echo.
echo === Step 1: Build RdpWrapKitHelper.exe ===
if not exist "%ROOT%scripts\RdpWrapKitHelper.cs" (
    echo ERROR: scripts\RdpWrapKitHelper.cs not found in %ROOT%scripts
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
    /out:"%ROOT%output\RdpWrapKitHelper.exe" "%ROOT%scripts\RdpWrapKitHelper.cs"
if errorlevel 1 (
    echo ERROR: RdpWrapKitHelper.exe build failed
    exit /b 1
)

REM Smoke test: the usage screen proves the new binary starts and its verb table
REM is wired up. It needs no privileges and writes nothing.
"%ROOT%output\RdpWrapKitHelper.exe" --help >nul 2>nul
if errorlevel 1 (
    echo ERROR: RdpWrapKitHelper.exe --help failed
    exit /b 1
)

if %SIGNONLY% equ 1 (
    echo.
    echo RdpWrapKitHelper.exe built successfully.
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
