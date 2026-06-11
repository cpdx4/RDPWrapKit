@echo off
setlocal EnableExtensions
REM ===========================================================================
REM RDPWrapKit Build Wrapper
REM ===========================================================================
REM
REM 1. Compiles RdpSignTool.exe from scripts\RdpSignTool.cs (C# source)
REM 2. Compiles RDPWrapKit-Setup.exe from RDPWrapKit.iss (Inno Setup)
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
if not exist "%ROOT%scripts\build_rdpcrypt.ps1" (
    echo ERROR: scripts\build_rdpcrypt.ps1 not found in %ROOT%scripts
    exit /b 1
)

powershell.exe -ExecutionPolicy Bypass -File "%ROOT%scripts\build_rdpcrypt.ps1"
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

REM Try to find ISCC.exe in common locations
set ISCC_EXE=
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set ISCC_EXE=C:\Program Files (x86)\Inno Setup 6\ISCC.exe
if not defined ISCC_EXE if exist "C:\Program Files\Inno Setup 6\ISCC.exe" set ISCC_EXE=C:\Program Files\Inno Setup 6\ISCC.exe
if not defined ISCC_EXE where ISCC.exe >nul 2>nul && set ISCC_EXE=ISCC.exe

if not defined ISCC_EXE (
    echo ERROR: Inno Setup 6 compiler not found.
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
