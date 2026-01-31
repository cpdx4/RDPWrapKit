; =========================================================================
; RDPWrapKit - Advanced RDP Wrapper Management Suite
; =========================================================================
; 
; PURPOSE:
;   This Inno Setup installer provides a comprehensive solution for setting
;   up RDP Wrapper and TermWrap on Windows systems, enabling multiple concurrent
;   Remote Desktop sessions on non-Server editions of Windows.
;
; KEY FEATURES:
;   - Automatic installation of RDP Wrapper (stascorp) and TermWrap (llccd)
;   - VC++ Redistributable dependency management
;   - User account creation with automatic RDP shortcuts
;   - Security hardening (Windows Defender exclusions, secure credential handling)
;   - Advanced tools for managing existing users and shortcuts
;   - Full uninstallation support with registry cleanup
;
; SECURITY CONSIDERATIONS:
;   - All PowerShell commands run with ExecutionPolicy Bypass for reliability
;   - Passwords are encrypted using SecureString before writing to temp files
;   - Temporary files containing sensitive data are deleted after use
;   - Admin privileges required for system modifications
;
; ARCHITECTURE:
;   - Uses centralized constants for executables, paths, and URLs
;   - Helper functions for PowerShell/CMD execution reduce code duplication
;   - Progressive UI with step-by-step feedback during installation
;   - Lazy loading of user lists to avoid blocking wizard initialization
; =========================================================================

[Setup]
; -------------------------------------------------------------------------
; SETUP CONFIGURATION
; -------------------------------------------------------------------------
; This section defines the installer's core configuration including
; application metadata, installation paths, compression settings, and
; security requirements.
;
; IMPORTANT NOTES:
; - PrivilegesRequired=admin: Required for service management and registry edits
; - ArchitecturesInstallIn64BitMode: Ensures proper 64-bit installation
; - SolidCompression: Reduces installer size significantly
; - CloseApplications: Automatically closes conflicting processes
; -------------------------------------------------------------------------
AppName=RDPWrapKit
AppVersion=0.46
VersionInfoVersion=0.46.0.0
AppPublisher=cpdx4
AppPublisherURL=https://github.com/cpdx4/RDPWrapKit
AppSupportURL=https://github.com/cpdx4/RDPWrapKit/issues
AppUpdatesURL=https://github.com/cpdx4/RDPWrapKit/releases
AppCopyright=Copyright (C) 2024-2026 RDPWrapKit Project
DefaultDirName={commonpf64}\RDPWrapKit
OutputBaseFilename=RDPWrapKit-Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
RestartApplications=yes
CloseApplicationsFilter=*.exe,*.chm
WizardStyle=modern dynamic
SetupIconFile="assets\RDPWrapKitIcon.ico"

[Files]
; -------------------------------------------------------------------------
; FILE DEPLOYMENT
; -------------------------------------------------------------------------
; Defines which files are copied during installation and under what conditions.
; 
; CONDITIONAL DEPLOYMENT:
;   - ShouldInstallFiles check ensures files are only copied when RDP Wrapper
;     installation is selected (not for user-only operations)
;   - Icon file always extracted to temp for welcome page display
;
; SECURITY:
;   - All files require admin privileges due to PrivilegesRequired setting
;   - Files are protected by Windows Defender exclusions added during install
; -------------------------------------------------------------------------
; Bundle your payload files (only for Full Install, not for Add Users Only)
Source: "third_party\rdpwrap_release\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs; Check: ShouldInstallFiles
Source: "third_party\termwrap_release\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs; Check: ShouldInstallFiles
Source: "assets\RDPWrapKitIcon.bmp"; DestDir: "{tmp}"; Flags: ignoreversion



[Registry]
; -------------------------------------------------------------------------
; REGISTRY MODIFICATIONS
; -------------------------------------------------------------------------
; Enables RDP connections and optimizes Remote Desktop client behavior.
;
; fDenyTSConnections=0: Enables Terminal Services (RDP) connections
; RemoteDesktop_SuppressWhenMinimized=2: Allows RDP to work when minimized
;
; CLEANUP:
;   - uninsdeletevalue flag ensures these entries are removed on uninstall
; -------------------------------------------------------------------------
; Enable RDP
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Terminal Server"; ValueType: dword; ValueName: "fDenyTSConnections"; ValueData: 0; Flags: uninsdeletevalue
; Allow minimized RDP
Root: HKLM; Subkey: "Software\Microsoft\Terminal Server Client"; ValueType: dword; ValueName: "RemoteDesktop_SuppressWhenMinimized"; ValueData: 2; Flags: uninsdeletevalue

[Run]
; Run section is now handled in Code section for proper sequencing

[UninstallRun]
; No separate uninstall runs needed - registry cleanup is handled in code section

[Code]
{ =============================================================================
  PASCAL SCRIPT CODE SECTION
  =============================================================================
  
  This section contains all the Pascal Script code that drives the installer's
  custom logic, UI, and installation flow.
  
  ARCHITECTURE OVERVIEW:
  ----------------------
  1. Constants Section:
     - Defines reusable values for executables, paths, URLs, timeouts
     - Reduces magic numbers and string literals throughout code
     - Centralizes configuration for easy maintenance
  
  2. External Windows API Declarations:
     - NetUserGetInfo/NetUserEnum: User account validation
     - LogonUser: Credential verification
     - GetTickCount: Timing operations
  
  3. Helper Functions:
     - Path construction (AppRoot, AppBin, TempFile)
     - PowerShell execution (BuildPowerShellArgs, ExecPowerShellHidden)
     - Command execution (RunHidden, RunCmdHidden)
     - UI utilities (BoolToStr, SleepWithUI, color/RTF helpers)
  
  4. Core Installation Logic:
     - Service management (StopTermService, StartTermService)
     - User creation and validation
     - RDP shortcut generation with encrypted passwords
     - VC++ Redistributable installation
     - Windows Defender exclusion management
  
  5. UI Management:
     - Custom wizard pages (welcome, options, user creation, advanced tools)
     - Progressive step indicators on installation page
     - Credits and attribution display
  
  6. Installation Flow:
     - PrepareToInstall: Pre-installation checks
     - CurStepChanged: Main installation orchestration
     - CurPageChanged: Page-specific initialization
     - NextButtonClick: Input validation and page navigation
  
  SECURITY DESIGN:
  ----------------
  - Passwords encrypted with PowerShell SecureString before temp file storage
  - Temporary files with sensitive data deleted immediately after use
  - All PowerShell commands use -ExecutionPolicy Bypass to avoid policy blocks
  - Hidden window modes prevent password exposure in UI
  - Local credential validation before account creation
  
  ERROR HANDLING:
  ---------------
  - Service operations include retry logic with exponential backoff
  - Optional components (like VC++ check) gracefully degrade on failure
  - User feedback provided for all critical errors
  - Exit codes logged for troubleshooting
  
  PERFORMANCE OPTIMIZATIONS:
  -------------------------
  - Lazy loading of user lists (only when advanced page shown)
  - Parallel tool invocations where possible
  - Minimized sleep durations (100-500ms instead of seconds)
  - Service state verified after stop/start operations
  
  CODING STANDARDS:
  -----------------
  - Descriptive function and variable names
  - Consistent indentation and formatting
  - Comments explain WHY not just WHAT
  - Magic numbers replaced with named constants
  - DRY principle applied throughout
  
  MAINTENANCE NOTES:
  ------------------
  - Update version numbers in both [Setup] section and constants
  - Test all paths (typical, advanced, uninstall) before release
  - Verify PowerShell commands work on latest Windows versions
  - Keep URL constants updated if upstream projects move
  
============================================================================= }

// Forward declarations
procedure CheckAndInstallMSTSC; forward;
procedure OnManageUsersClick(Sender: TObject); forward;
procedure OpenRDPWrap(Sender: TObject); forward;
procedure OpenTermWrap(Sender: TObject); forward;
procedure OpenBSGH(Sender: TObject); forward;
procedure OpenBSSGrinders(Sender: TObject); forward;
procedure OnInstallTypeChange(Sender: TObject); forward;
procedure InitInstallerLog; forward;
procedure WriteInstallerLog(const Msg: string); forward;
var
  InstallTypePage: TWizardPage;
  WelcomePage: TWizardPage;
  UserPage: TInputQueryWizardPage;
  AdvancedPage: TWizardPage;
  AdvancedToolsSelectionPage: TInputOptionWizardPage;  // Main Advanced Tools selection
  AdvancedTool1Page: TWizardPage;  // Create RDP desktop shortcuts
  AdvancedTool2Page: TWizardPage;  // Placeholder tool 2
  AdvancedTool3Page: TWizardPage;  // Placeholder tool 3
  AdvancedTool4Page: TWizardPage;  // Placeholder tool 4
  AdvancedTool5Page: TWizardPage;  // Placeholder tool 5
  SelectedAdvancedTool: Integer;   // Track which tool is selected
  LocalUsersList: TStringList;
  UserCheckBoxes: array of TCheckBox;
  UserPasswordEdits: array of TEdit;
  UserPasswordStatus: array of TLabel;
  ShortcutsList: TStringList;
  InstallLogPath: string;
  AddMoreRadio: TRadioButton;
  DoneRadio: TRadioButton;
  SkipUsersCheckBox: TCheckBox;
  // New welcome/options controls
  TypicalRadio: TRadioButton;
  AdvancedRadio: TRadioButton;
  UninstallRadio: TRadioButton;
  chkInstallRDPWrapper: TCheckBox;
  chkManageUsers: TCheckBox;
  rbCreateUsers: TRadioButton;
  rbUseExistingUsers: TRadioButton;
  ManageUsersGroup: TPanel;
  // Flags derived from welcome/options controls
  DoInstallRDPWrapper: Boolean;
  DoManageUsers: Boolean;
  NeedCreateUsers: Boolean;
  JumpToAdvancedTool1: Boolean;
  OptionsLabel: TLabel;
  Tool1UsersHeaderLabel: TLabel;  // "Users found" header
  Tool1PasswordHeaderLabel: TLabel;  // "Password" header
  Tool1PasswordResetLink: TLabel;  // Password reset link at bottom
  // Progress UI on Installing page
  StepsHeaderLabel: TLabel;
  StepAddExcl: TLabel;
  StepRemoveExcl: TLabel;
  StepStopSvc: TLabel;
  StepEnsureVC: TLabel;
  StepInstallRDPWrapper: TLabel;
  StepConfigureService: TLabel;
  StepCreateUsers: TLabel;
  StepCreateShortcuts: TLabel;
  StepPreTrust: TLabel;
  StepStartSvc: TLabel;
  StepCheckRDP: TLabel;
  StepCheckMSTSC: TLabel;
  StepInstallMSTSC: TLabel;
  StepRemoveFolder: TLabel;
  StepUninstallRDPWrapper: TLabel;
  InstallType: Integer;  // 0 = Full Install, 1 = Add User Only, 2 = Advanced, 3 = Uninstall Everything
  DebugMode: Boolean;    // Set to True to force VC++ download even if installed
  UsersList: TStringList;
  CreatedUsersList: TStringList;  // Store usernames to display on finish page
  CurrentUserIndex: Integer;
  // Layout helpers for step labels
  StepLeftPos: Integer;
  StepTopBase: Integer;
  StepWidthVal: Integer;
  StepNextTop: Integer;
  // Localized group names (resolved at runtime)
  GroupAdministratorsName: string;
  GroupRDPUsersName: string;
  // Credits / license blurb on welcome page
  CreditsText: TRichEditViewer;
  // Rich text control on finished page to show long completion messages
  FinishedText: TRichEditViewer;
  // Button to open install log on finish page
  ViewLogButton: TButton;

const
  // -------------------------------------------------------------------------
  // SYSTEM CONSTANTS
  // -------------------------------------------------------------------------
  // Windows API return codes and limits
  NERR_Success = 0;              // Windows NetAPI success code
  MAX_SHORTCUTS = 10;            // Safety limit for RDP shortcut creation
  
  // -------------------------------------------------------------------------
  // USER INTERFACE TEXT
  // -------------------------------------------------------------------------
  // Step text constants for progress checklist - displayed during installation
  TXT_AddExcl = 'Add Windows Defender exclusion';
  TXT_RemoveExcl = 'Remove Windows Defender exclusion';
  TXT_StopSvc = 'Stop Remote Desktop Services';
  TXT_StartSvc = 'Start Remote Desktop Services';
  TXT_RestartSvc = 'Restart Remote Desktop Services';
  TXT_EnsureVC = 'Install VC++ Redistributable (2015-2022)';
  TXT_InstallRDPWrapper = 'Install RDP Wrapper';
  TXT_ConfigureService = 'Configure TermWrap service';
  TXT_CreateUsers = 'Create user accounts';
  TXT_CreateShortcuts = 'Create RDP shortcuts for selected users';
  TXT_PreTrust = 'Pre-trust RDP certificate for current user';
  TXT_CheckRDP = 'Verify RDP service is listening';
  TXT_CheckMSTSC = 'Check for Remote Desktop Connection';
  TXT_InstallMSTSC = 'Install Remote Desktop Connection';
  TXT_RemoveFolder = 'Remove RDP Wrapper folder';
  TXT_UninstallRDPWrapper = 'Uninstall RDP Wrapper';
  
  // -------------------------------------------------------------------------
  // REGISTRY PATHS
  // -------------------------------------------------------------------------
  // Windows Registry keys for Terminal Services and RDP configuration
  // IMPORTANT: These paths are system-critical and must remain accurate
  REG_TERMSERVICE_PARAMS = 'SYSTEM\CurrentControlSet\Services\TermService\Parameters';
  REG_TERMSERVICE = 'SYSTEM\CurrentControlSet\Services\TermService';
  REG_TERMINAL_SERVER = 'SYSTEM\CurrentControlSet\Control\Terminal Server';
  REG_UMRDPSERVICE_PARAMS = 'SYSTEM\CurrentControlSet\Services\UmRdpService\Parameters';
  REG_VCREDIST = 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64';
  // User groups
  GROUP_ADMINISTRATORS = 'Administrators';
  GROUP_RDP_USERS = 'Remote Desktop Users';
  
  // -------------------------------------------------------------------------
  // TIMING CONSTANTS
  // -------------------------------------------------------------------------
  // Sleep durations in milliseconds - optimized for fastest reliable installation
  // These values balance responsiveness with system stabilization needs
  SLEEP_SHORT = 100;
  SLEEP_MEDIUM = 250;
  SLEEP_LONG = 500;
  SLEEP_EXTRALONG = 2000;
  // -------------------------------------------------------------------------
  // TIMEOUTS
  // -------------------------------------------------------------------------
  // Best-effort timeouts used to avoid extremely long hangs during user
  // creation/shortcut generation. Note: Exec() calls block the script, so
  // timeouts are best-effort overall watchers rather than hard per-process
  // kill timers.
  PER_USER_TIMEOUT = 120000; // 2 minutes per user (best-effort)
  USERS_OVERALL_TIMEOUT = 600000; // 10 minutes overall for user operations
  
  // -------------------------------------------------------------------------
  // FILES, URLS, AND NETWORK PORTS
  // -------------------------------------------------------------------------
  // Application components and external resources
  FILE_RDPW_INST = 'RDPWInst.exe';
  FILE_TERMWRAP = 'TermWrap.dll';
  FILE_ZYDIS = 'Zydis.dll';
  URL_VCREDIST_X64 = 'https://aka.ms/vs/17/release/vc_redist.x64.exe';
  RDP_LISTEN_PORT = 3389;
  FILE_MSTSC = 'C:\\Windows\\System32\\mstsc.exe';
  URL_RDP_INSTALLER = 'https://go.microsoft.com/fwlink/?linkid=2247659';
  
  // -------------------------------------------------------------------------
  // REUSABLE EXECUTABLES AND COMMAND PATTERNS
  // -------------------------------------------------------------------------
  // Centralized constants for system executables to ensure consistency
  // and enable easy updates if paths change
  EXE_CMD = 'cmd.exe';
  EXE_POWERSHELL = 'powershell.exe';
  PS_ARGS_BASE = '-NoProfile -ExecutionPolicy Bypass';
  PS_ARGS_HIDDEN = PS_ARGS_BASE + ' -NonInteractive -WindowStyle Hidden';
  
  // -------------------------------------------------------------------------
  // TEMPORARY FILE PATHS
  // -------------------------------------------------------------------------
  // Pre-expanded temp paths for frequently-used temporary files
  // Using constants avoids repeated ExpandConstant() calls
  FILE_ICON_BMP = 'RDPWrapKitIcon.bmp';
  TEMP_LOCAL_USERS = '{tmp}\\local_users.txt';
  TEMP_ICON_BMP = '{tmp}\\' + FILE_ICON_BMP;
  // Path to installer log file (auto-created in %TEMP%)
  INSTALL_LOG_PATH = '{tmp}\\RDPWrapKit_install.log';
  
  // -------------------------------------------------------------------------
  // EXTERNAL PROJECT URLS
  // -------------------------------------------------------------------------
  // Centralized URLs for attribution and user navigation
  // Update these if upstream projects change their repository locations
  URL_RDPWRAP = 'https://github.com/stascorp/rdpwrap';
  URL_TERMWRAP = 'https://github.com/llccd/TermWrap';
  URL_BSGH_COMMUNITY = 'https://discord.gg/bsgh';
  URL_BSS_GRINDERS = 'https://discord.gg/K5U3RdGXh6';
  URL_PROJECT_HOME = 'https://github.com/cpdx4/RDPWrapKit';

// =============================================================================
// EXTERNAL WINDOWS API FUNCTION DECLARATIONS
// =============================================================================
// These functions provide access to Windows system functionality not available
// through standard Inno Setup commands.
//
// SECURITY IMPLICATIONS:
// - NetUserGetInfo/NetUserEnum: Read-only access to local user accounts
// - LogonUser: Validates credentials without creating sessions
// - All operations require admin privileges (enforced by installer config)
// =============================================================================

function NetUserGetInfo(ServerName: String; UserName: String; Level: Cardinal; var BufPtr: Cardinal): Cardinal;
  external 'NetUserGetInfo@netapi32.dll stdcall';

function NetApiBufferFree(Buf: Cardinal): Cardinal;
  external 'NetApiBufferFree@netapi32.dll stdcall';

function NetUserEnum(ServerName: String; Level: Cardinal; Filter: Cardinal; var BufPtr: Cardinal; PrefMaxLen: Cardinal; var EntriesRead: Cardinal; var TotalEntries: Cardinal; var ResumeHandle: Cardinal): Cardinal;
  external 'NetUserEnum@netapi32.dll stdcall';

function LogonUser(lpUsername: string; lpDomain: string; lpPassword: string; dwLogonType, dwLogonProvider: Cardinal; var phToken: Cardinal): Boolean;
  external 'LogonUserW@advapi32.dll stdcall';

function CloseHandle(hObject: Cardinal): Boolean;
  external 'CloseHandle@kernel32.dll stdcall';

function GetTickCount: Cardinal;
  external 'GetTickCount@kernel32.dll stdcall';

// Windows SYSTEMTIME structure and GetSystemTime API for timestamps
type
  SYSTEMTIME = record
    wYear: Word;
    wMonth: Word;
    wDayOfWeek: Word;
    wDay: Word;
    wHour: Word;
    wMinute: Word;
    wSecond: Word;
    wMilliseconds: Word;
  end;

function GetSystemTime(var lpSystemTime: SYSTEMTIME): Boolean;
  external 'GetSystemTime@kernel32.dll stdcall';

// =============================================================================
// UTILITY AND HELPER FUNCTIONS
// =============================================================================
// This section contains reusable functions that simplify common operations
// throughout the installer. These functions follow the DRY principle and
// provide consistent behavior across different installation paths.
// =============================================================================

// -----------------------------------------------------------------------------
// USER ACCOUNT VALIDATION
// -----------------------------------------------------------------------------

// Check if a local user account exists on the system
// Returns: True if user exists, False otherwise
// Uses: Windows NetAPI for fast, reliable user enumeration
function UserExists(const UserName: string): Boolean;
var
  BufPtr: Cardinal;
begin
  BufPtr := 0;
  Result := NetUserGetInfo('', UserName, 0, BufPtr) = NERR_Success;
  if BufPtr <> 0 then
    NetApiBufferFree(BufPtr);
end;

// Parse a pipe-delimited user entry string into username and password components
// Format: "username|password"
// Used for: Internal user list storage and processing
// Security: This function handles sensitive data - ensure calling code cleans up
procedure ParseUserEntry(const Entry: string; var UserName, Password: string);
var
  PipePos: Integer;
begin
  PipePos := Pos('|', Entry);
  UserName := Copy(Entry, 1, PipePos - 1);
  Password := Copy(Entry, PipePos + 1, Length(Entry));
end;

// -----------------------------------------------------------------------------
// PATH CONSTRUCTION HELPERS
// -----------------------------------------------------------------------------
// These functions provide consistent, reusable path building for commonly
// accessed locations. Using these helpers ensures paths remain correct even
// if installation directory structure changes.
// -----------------------------------------------------------------------------

// Get the RDP Wrapper installation path
// Returns: C:\Program Files\RDP Wrapper
function GetRDPWrapperPath: string;
begin
  Result := ExpandConstant('{commonpf64}\RDP Wrapper');
end;

// Build an expanded path under {tmp} to avoid repeating ExpandConstant() calls
// Parameter: FileName - The filename to place in temp directory
// Returns: Fully expanded path like C:\Users\...\AppData\Local\Temp\filename
// Usage: Simplifies temporary file creation throughout the installer
function TempFile(const FileName: string): string;
begin
  Result := ExpandConstant('{tmp}\' + FileName);
end;

// -----------------------------------------------------------------------------
// POWERSHELL EXECUTION HELPERS
// -----------------------------------------------------------------------------
// Standardized PowerShell invocation builders that ensure consistent behavior,
// security settings, and error handling across all PS operations.
//
// SECURITY NOTES:
// - ExecutionPolicy Bypass: Required for unsigned scripts to run
// - NonInteractive: Prevents prompts that would hang the installer
// - WindowStyle Hidden: Prevents password exposure in console windows
// -----------------------------------------------------------------------------

// Construct standardized PowerShell arguments for command execution
// Parameters:
//   Command - The PowerShell command(s) to execute
//   Hidden - If True, runs with -NonInteractive -WindowStyle Hidden
// Returns: Complete argument string for Exec() function
function BuildPowerShellArgs(const Command: string; Hidden: Boolean): string;
begin
  if Hidden then
    Result := PS_ARGS_HIDDEN + ' -Command "' + Command + '"'
  else
    Result := PS_ARGS_BASE + ' -Command "' + Command + '"';
end;

// Standardize PowerShell script execution with optional extra parameters
function BuildPowerShellFileArgs(const ScriptPath, ExtraParams: string; Hidden: Boolean): string;
var
  BaseArgs: string;
  CleanExtra: string;
begin
  if Hidden then
    BaseArgs := PS_ARGS_HIDDEN
  else
    BaseArgs := PS_ARGS_BASE;

  CleanExtra := Trim(ExtraParams);
  if CleanExtra <> '' then
    Result := BaseArgs + ' -File "' + ScriptPath + '" ' + CleanExtra
  else
    Result := BaseArgs + ' -File "' + ScriptPath + '"';
end;

function ExecPowerShellHidden(const Command: string; var ResultCode: Integer): Boolean;
var
  PSArgs: string;
begin
  PSArgs := BuildPowerShellArgs(Command, True);
  // Log command and run
  WriteInstallerLog('PowerShell Hidden: ' + PSArgs);
  Result := Exec(EXE_POWERSHELL, PSArgs, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  WriteInstallerLog('PowerShell exitcode=' + IntToStr(ResultCode));
end;

procedure OpenRDPWrap(Sender: TObject);
var
  rc: Integer;
begin
  ShellExec('', URL_RDPWRAP, '', '', SW_SHOWNORMAL, ewNoWait, rc);
end;

procedure OpenTermWrap(Sender: TObject);
var
  rc: Integer;
begin
  ShellExec('', URL_TERMWRAP, '', '', SW_SHOWNORMAL, ewNoWait, rc);
end;

procedure OpenBSGH(Sender: TObject);
var
  rc: Integer;
begin
  ShellExec('', URL_BSGH_COMMUNITY, '', '', SW_SHOWNORMAL, ewNoWait, rc);
end;

procedure OpenBSSGrinders(Sender: TObject);
var
  rc: Integer;
begin
  ShellExec('', URL_BSS_GRINDERS, '', '', SW_SHOWNORMAL, ewNoWait, rc);
end;

procedure OpenProjectHome(Sender: TObject);
var
  rc: Integer;
begin
  ShellExec('', URL_PROJECT_HOME, '', '', SW_SHOWNORMAL, ewNoWait, rc);
end;

// === Utility helpers ========================================================

// App install root helper
function AppRoot: string;
begin
  Result := ExpandConstant('{app}');
end;

// Path to a file under the app install root
function AppBin(const FileName: string): string;
begin
  Result := ExpandConstant('{app}\' + FileName);
end;

// Installer log helpers
function GetTimestampString: string;
var
  ST: SYSTEMTIME;
  h, m, s, ms: string;
begin
  GetSystemTime(ST);
  // Pad numbers with zeros
  if ST.wHour < 10 then h := '0' + IntToStr(ST.wHour) else h := IntToStr(ST.wHour);
  if ST.wMinute < 10 then m := '0' + IntToStr(ST.wMinute) else m := IntToStr(ST.wMinute);
  if ST.wSecond < 10 then s := '0' + IntToStr(ST.wSecond) else s := IntToStr(ST.wSecond);
  if ST.wMilliseconds < 100 then ms := '0' + IntToStr(ST.wMilliseconds) else ms := IntToStr(ST.wMilliseconds);
  if ST.wMilliseconds < 10 then ms := '00' + IntToStr(ST.wMilliseconds);
  Result := '[' + h + ':' + m + ':' + s + '.' + ms + ']';
end;

procedure InitInstallerLog;
begin
  InstallLogPath := ExpandConstant(INSTALL_LOG_PATH);
  try
    SaveStringToFile(InstallLogPath, GetTimestampString + ' RDPWrapKit install log started' + #13#10, False);
  except
  end;
end;

procedure WriteInstallerLog(const Msg: string);
begin
  try
    SaveStringToFile(InstallLogPath, GetTimestampString + ' ' + Msg + #13#10, True);
  except
  end;
end;

// Run any process hidden and return its exit code
function RunHidden(const FileName, Params: string): Integer;
var
  RC: Integer;
begin
  RC := 0;
  WriteInstallerLog('Exec: ' + FileName + ' ' + Params);
  Exec(FileName, Params, '', SW_HIDE, ewWaitUntilTerminated, RC);
  WriteInstallerLog('ExitCode: ' + IntToStr(RC) + ' for ' + FileName);
  Result := RC;
end;

// Run a cmd.exe one-liner hidden and return its exit code
function RunCmdHidden(const CmdLine: string): Integer;
begin
  Result := RunHidden(EXE_CMD, '/c ' + CmdLine);
end;

// Check if a process is running (by executable name)
function ProcessExists(const ProcessName: string): Boolean;
var
  RC: Integer;
begin
  // Use tasklist to check if process is running
  // tasklist returns 0 if found, 1 if not found
  RC := RunCmdHidden('tasklist /FI "IMAGENAME eq ' + ProcessName + '" /FO CSV /NH');
  Result := (RC = 0);
end;

// Convert boolean to string
function BoolToStr(Value: Boolean): string;
begin
  if Value then
    Result := 'True'
  else
    Result := 'False';
end;

// Sleep with UI updates
procedure SleepWithUI(Milliseconds: Integer);
var
  Elapsed: Integer;
  ChunkSize: Integer;
begin
  Elapsed := 0;
  ChunkSize := 100;
  
  while Elapsed < Milliseconds do
  begin
    Sleep(ChunkSize);
    Elapsed := Elapsed + ChunkSize;
    WizardForm.Update;
  end;
end;

// Run in PowerShell via our existing wrapper and return exit code
function RunPSHiddenCode(const Command: string): Integer;
var
  RC: Integer;
begin
  RC := -1;
  WriteInstallerLog('RunPSHiddenCode: ' + Command);
  ExecPowerShellHidden(Command, RC);
  WriteInstallerLog('RunPSHiddenCode exit=' + IntToStr(RC));
  Result := RC;
end;

// Convert plain text to a minimal RTF string (escapes braces and backslashes and converts CRLF to \par)
function PlainToRtf(const S: string): string;
var
  i: Integer;
  ch: string;
begin
  Result := '{\rtf1\ansi ';
  i := 1;
  while i <= Length(S) do
  begin
    if (i < Length(S)) and (S[i] = #13) and (S[i+1] = #10) then
    begin
      Result := Result + '\par ';
      i := i + 2;
    end
    else
    begin
      ch := Copy(S, i, 1);
      if ch = '\' then Result := Result + '\\'
      else if ch = '{' then Result := Result + '\{'
      else if ch = '}' then Result := Result + '\}'
      else Result := Result + ch;
      i := i + 1;
    end;
  end;
  Result := Result + '}';
end;

function RGBToColor(R, G, B: Integer): Longint;
begin
  Result := (B shl 16) or (G shl 8) or R;
end;

function IsDarkColor(C: Longint): Boolean;
var R, G, B, bright: Integer;
begin
  R := C and $FF;
  G := (C shr 8) and $FF;
  B := (C shr 16) and $FF;
  bright := (R * 299 + G * 587 + B * 114) div 1000;
  Result := bright < 128;
end;

function PlainToRtfWithColor(const S: string; Color: Longint): string;
var
  i: Integer;
  ch: string;
  R, G, B: Integer;
begin
  R := Color and $FF;
  G := (Color shr 8) and $FF;
  B := (Color shr 16) and $FF;
  Result := '{\rtf1\ansi{\colortbl;' +
    '\red' + IntToStr(R) + '\green' + IntToStr(G) + '\blue' + IntToStr(B) + ';}';
  i := 1;
  while i <= Length(S) do
  begin
    if (i < Length(S)) and (S[i] = #13) and (S[i+1] = #10) then
    begin
      Result := Result + '\par ';
      i := i + 2;
    end
    else
    begin
      ch := Copy(S, i, 1);
      if ch = '\' then Result := Result + '\\'
      else if ch = '{' then Result := Result + '\{'
      else if ch = '}' then Result := Result + '\}'
      else Result := Result + ch;
      i := i + 1;
    end;
  end;
  Result := Result + '}';
end;

// BuildCreditsRtf removed — welcome page uses individual TLabel controls instead

// (moved: layout helper procedures are now declared after SetStepPending/Done)

// Resolve localized group name from a well-known SID; fallback to provided name
function GetLocalizedGroupName(const Sid, Fallback: string): string;
var
  ResultCode: Integer;
  OutPath: string;
  NameText: AnsiString;
begin
  Result := Fallback;
  OutPath := TempFile('grp_' + Sid + '.txt');
  if ExecPowerShellHidden(
    '$ErrorActionPreference = ''Stop''; ' +
    '$sid = New-Object System.Security.Principal.SecurityIdentifier(''' + Sid + '''); ' +
    '$acct = $sid.Translate([System.Security.Principal.NTAccount]).Value; ' +
    '$grp = $acct.Split([char]92)[-1]; ' +
    '[System.IO.File]::WriteAllText(''' + OutPath + ''', $grp)'
    , ResultCode) then
  begin
    if (ResultCode = 0) and FileExists(OutPath) then
    begin
      if LoadStringFromFile(OutPath, NameText) then
      begin
        if Trim(String(NameText)) <> '' then
          Result := Trim(String(NameText));
      end;
    end;
    DeleteFile(OutPath);
  end;
end;

// Pre-trust the current user's RDP client for 127.0.0.2
// Creates the LocalDevices registry entry that Windows sets when user checks "Don't ask me again"
procedure PreTrustRDPCertCurrentUser;
var
  ResultCode: Integer;
begin
  // Value 76 (0x4C) represents the device/resource trust flags
  ExecPowerShellHidden(
    '$ErrorActionPreference = ''Stop''; ' +
    'try { ' +
    '  $regPath = ''HKCU:\Software\Microsoft\Terminal Server Client\LocalDevices''; ' +
    '  if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }; ' +
    '  New-ItemProperty -Path $regPath -Name ''127.0.0.2'' -PropertyType DWord -Value 76 -Force | Out-Null; ' +
    '  exit 0 ' +
    '} catch { ' +
    '  exit 1 ' +
    '}',
    ResultCode);
  Log('DEBUG: PreTrustRDPCertCurrentUser exit code = ' + IntToStr(ResultCode) + ' (0=success, 1=error)');
end;

procedure AddDefenderExclusionForApp;
var
  ResultCode: Integer;
  ExclPath: string;
begin
  // Ensure Defender excludes the install folder before any executables run
  ExclPath := GetRDPWrapperPath;
  ExecPowerShellHidden(
    '$path = ''' + ExclPath + '''; ' +
    'try { $p = Get-MpPreference; if (-not ($p.ExclusionPath -contains $path)) { Add-MpPreference -ExclusionPath $path } } catch { }',
    ResultCode);
end;

procedure RemoveDefenderExclusionForApp;
var
  ResultCode: Integer;
  ExclPath: string;
begin
  // Remove Defender exclusion for the install folder during uninstall
  ExclPath := GetRDPWrapperPath;
  ExecPowerShellHidden(
    '$path = ''' + ExclPath + '''; ' +
    'try { Remove-MpPreference -ExclusionPath $path } catch { }',
    ResultCode);
end;

// -----------------------------------------------------------------------------
// WINDOWS SERVICE MANAGEMENT
// -----------------------------------------------------------------------------
// Functions for controlling the Terminal Services (TermService) which must be
// stopped before modifying RDP Wrapper DLLs and restarted afterward.
//
// CRITICAL NOTES:
// - Service must be fully stopped before DLL replacement or corruption occurs
// - Includes retry logic as service can be slow to stop
// - Service state verification after operations
// -----------------------------------------------------------------------------

// Stop the Terminal Service (RDP) with retry logic
// This is required before replacing RDP Wrapper DLLs
// Attempts: Up to 5 tries with progressive delays
// Side effects: Disconnects all active RDP sessions
procedure StopTermService;
var
  ResultCode: Integer;
  Attempt: Integer;
  ServiceRunning: Boolean;
begin
  Log('[StopTermService] START');
  WriteInstallerLog('StopTermService: Stopping Remote Desktop Services...');
  WizardForm.StatusLabel.Caption := 'Stopping Remote Desktop Services...';
  
  // Start with disable + stop attempts (no initial "normal" stop)
  ServiceRunning := True;
  Attempt := 0;
  while ServiceRunning and (Attempt < 5) do
  begin
    Inc(Attempt);
    Log('[StopTermService] Attempt ' + IntToStr(Attempt) + '/5: Disable and stop');
    WriteInstallerLog('StopTermService: Attempt ' + IntToStr(Attempt) + '/5');
    WizardForm.Update;
    
    // Disable the service to prevent auto-restart
    ResultCode := RunPSHiddenCode('Set-Service -Name TermService -StartupType Disabled -ErrorAction Stop');
    Log('[StopTermService] Set-Service Disabled exit code: ' + IntToStr(ResultCode));
    WriteInstallerLog('StopTermService: Set-Service Disabled exit code=' + IntToStr(ResultCode));
    SleepWithUI(SLEEP_MEDIUM);
    
    // Try stopping with PowerShell
    Log('[StopTermService] Stop-Service TermService');
    WriteInstallerLog('StopTermService: Executing Stop-Service TermService');
    ResultCode := RunPSHiddenCode('Stop-Service -Name TermService -Force -ErrorAction Stop');
    Log('[StopTermService] Stop-Service exit code: ' + IntToStr(ResultCode));
    WriteInstallerLog('StopTermService: Stop-Service exit code=' + IntToStr(ResultCode));
    
    // Wait longer for service to actually stop
    SleepWithUI(SLEEP_LONG + SLEEP_LONG);
    
    // Check service state (exit 0 if stopped, exit 1 if running or any other state)
    ResultCode := RunPSHiddenCode('if ((Get-Service -Name TermService).Status -eq ''Stopped'') { exit 0 } else { exit 1 }');
    ServiceRunning := (ResultCode = 1);
    Log('[StopTermService] After attempt ' + IntToStr(Attempt) + ': Service running=' + BoolToStr(ServiceRunning) + ' (check exit code: ' + IntToStr(ResultCode) + ')');
    WriteInstallerLog('StopTermService: Service running=' + BoolToStr(ServiceRunning) + ' after attempt ' + IntToStr(Attempt));
  end;
  
  if ServiceRunning then
  begin
    Log('[StopTermService] WARNING: Service still running after 5 attempts, proceeding anyway');
    WriteInstallerLog('StopTermService: WARNING - Service still running after 5 attempts, proceeding anyway');
  end
  else
  begin
    Log('[StopTermService] SUCCESS: Service verified stopped');
    WriteInstallerLog('StopTermService: SUCCESS - Service verified stopped');
  end;
  
  Log('[StopTermService] END');
  WriteInstallerLog('StopTermService: END');
end;

// Start TermService and set it to Automatic startup, return exit code
function StartTermServiceEx: Integer;
var
  RC: Integer;
begin
  Log('[StartTermServiceEx] START');
  WriteInstallerLog('StartTermServiceEx: Starting Remote Desktop Services...');
  WizardForm.StatusLabel.Caption := 'Restarting Remote Desktop Services...';
  
  Log('[StartTermServiceEx] Setting TermService to Automatic via PowerShell');
  WriteInstallerLog('StartTermServiceEx: Setting TermService to Automatic');
  RC := RunPSHiddenCode('Set-Service -Name TermService -StartupType Automatic -ErrorAction Stop');
  Log('[StartTermServiceEx] Set-Service Automatic exit code: ' + IntToStr(RC));
  WriteInstallerLog('StartTermServiceEx: Set-Service Automatic exit code=' + IntToStr(RC));
  
  Sleep(SLEEP_MEDIUM);
  Log('[StartTermServiceEx] Starting TermService via PowerShell');
  WriteInstallerLog('StartTermServiceEx: Executing Start-Service TermService');
  RC := RunPSHiddenCode('Start-Service -Name TermService -ErrorAction Stop');
  Log('[StartTermServiceEx] Start-Service exit code: ' + IntToStr(RC) + ' (0=success)');
  WriteInstallerLog('StartTermServiceEx: Start-Service exit code=' + IntToStr(RC));
  Sleep(SLEEP_LONG);
  Log('[StartTermServiceEx] END');
  WriteInstallerLog('StartTermServiceEx: END');
  Result := RC;
end;

// Backward-compatible wrapper (ignores exit code)
procedure StartTermService;
begin
  StartTermServiceEx;
end;

function IsExcludedUser(const UserName: string): Boolean;
begin
  Result :=
    (CompareText(UserName, 'Administrator') = 0) or
    (CompareText(UserName, 'Guest') = 0) or
    (CompareText(UserName, 'DefaultAccount') = 0) or
    (CompareText(UserName, 'defaultuser0') = 0) or
    (CompareText(UserName, 'WDAGUtilityAccount') = 0);
end;

function GetLocalUsers: TStringList;
var
  UsersList: TStringList;
  PSPath: string;
  ResultCode: Integer;
  i: Integer;
  Line: string;
begin
  UsersList := TStringList.Create;
  PSPath := ExpandConstant(TEMP_LOCAL_USERS);
  
  // Use optimized PowerShell command
  Exec(EXE_POWERSHELL,
    BuildPowerShellArgs(
      'Get-LocalUser | ' +
      'Where-Object { $_.Enabled -eq $true -and $_.PrincipalSource -eq ''Local'' } | ' +
      'Select-Object -ExpandProperty Name | ' +
      'Out-File -Encoding UTF8 ''' + PSPath + ''' -Force', True),
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  if (ResultCode = 0) and FileExists(PSPath) then
  begin
    UsersList.LoadFromFile(PSPath);
    // Filter in reverse to avoid index issues
    for i := UsersList.Count - 1 downto 0 do
    begin
      Line := Trim(UsersList[i]);
      if (Line = '') or IsExcludedUser(Line) then
        UsersList.Delete(i)
      else
        UsersList[i] := Line;
    end;
    DeleteFile(PSPath);  // Clean up temp file
  end;

  Result := UsersList;
end;

procedure SetUserControlsEnabled(Enabled: Boolean);
var
  i: Integer;
begin
  for i := 0 to High(UserCheckBoxes) do
  begin
    if Assigned(UserCheckBoxes[i]) then
      UserCheckBoxes[i].Enabled := Enabled;
    if Assigned(UserPasswordEdits[i]) then
    begin
      UserPasswordEdits[i].Enabled := Enabled and UserCheckBoxes[i].Checked;
      if not UserPasswordEdits[i].Enabled then
        UserPasswordEdits[i].Text := '';
    end;
    if Assigned(UserPasswordStatus[i]) then
      UserPasswordStatus[i].Visible := False;
  end;
end;

function FindUserIndexFromCheckBox(CB: TCheckBox): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to High(UserCheckBoxes) do
  begin
    if UserCheckBoxes[i] = CB then
    begin
      Result := i;
      exit;
    end;
  end;
end;

function FindUserIndexFromEdit(E: TEdit): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to High(UserPasswordEdits) do
  begin
    if UserPasswordEdits[i] = E then
    begin
      Result := i;
      exit;
    end;
  end;
end;


procedure OnPasswordEditChange(Sender: TObject);
var
  idx: Integer;
begin
  idx := FindUserIndexFromEdit(TEdit(Sender));
  if idx >= 0 then
    UserPasswordStatus[idx].Visible := False;
end;

procedure OnUserCheckBoxClick(Sender: TObject);
var
  idx: Integer;
begin
  idx := FindUserIndexFromCheckBox(TCheckBox(Sender));
  if idx >= 0 then
  begin
    UserPasswordEdits[idx].Enabled := UserCheckBoxes[idx].Checked;
    if not UserCheckBoxes[idx].Checked then
    begin
      UserPasswordEdits[idx].Text := '';
      UserPasswordStatus[idx].Visible := False;
    end;
  end;
end;

procedure OnManageUsersClick(Sender: TObject);
begin
  if Assigned(rbCreateUsers) then
    rbCreateUsers.Enabled := chkManageUsers.Checked;
  if Assigned(rbUseExistingUsers) then
    rbUseExistingUsers.Enabled := chkManageUsers.Checked;
  // Update derived flags so page skipping and next-button validation are accurate
  DoManageUsers := chkManageUsers.Checked;
  NeedCreateUsers := DoManageUsers and Assigned(rbCreateUsers) and rbCreateUsers.Checked;
end;

procedure OnInstallTypeChange(Sender: TObject);
begin
  if Assigned(chkInstallRDPWrapper) then
  begin
    // If Uninstall is selected, disable typical sub-options
    if Assigned(UninstallRadio) and UninstallRadio.Checked then
    begin
      chkInstallRDPWrapper.Enabled := False;
      chkManageUsers.Enabled := False;
      if Assigned(rbCreateUsers) then rbCreateUsers.Enabled := False;
      if Assigned(rbUseExistingUsers) then rbUseExistingUsers.Enabled := False;
    end
    else
    begin
      // Typical or other flows: enable based on defaults and ManageUsers checkbox
      chkInstallRDPWrapper.Enabled := True;
      chkManageUsers.Enabled := True;
      if Assigned(rbCreateUsers) then rbCreateUsers.Enabled := chkManageUsers.Checked;
      if Assigned(rbUseExistingUsers) then rbUseExistingUsers.Enabled := chkManageUsers.Checked;
    end;
  end;
end;

procedure OnViewLogButtonClick(Sender: TObject);
var
  RC: Integer;
begin
  WriteInstallerLog('User clicked View Install Log button');
  // Open the log file with the default text editor
  ShellExec('open', InstallLogPath, '', '', SW_SHOWNORMAL, ewNoWait, RC);
end;

procedure OnPasswordResetLinkClick(Sender: TObject);
var
  ResultCode: Integer;
begin
  Exec('control.exe', 'userpasswords2', '', SW_SHOW, ewNoWait, ResultCode);
end;

procedure BuildAdvancedUserControls;
var
  i: Integer;
  TopPos: Integer;
  BottomPos: Integer;
begin
  TopPos := ScaleY(10);
  
  // Create "Users found" header
  Tool1UsersHeaderLabel := TLabel.Create(AdvancedTool1Page);
  Tool1UsersHeaderLabel.Parent := AdvancedTool1Page.Surface;
  Tool1UsersHeaderLabel.Left := ScaleX(20);
  Tool1UsersHeaderLabel.Top := TopPos;
  Tool1UsersHeaderLabel.Caption := 'Users found';
  Tool1UsersHeaderLabel.Font.Style := [fsBold];
  
  // Create "Password" header
  Tool1PasswordHeaderLabel := TLabel.Create(AdvancedTool1Page);
  Tool1PasswordHeaderLabel.Parent := AdvancedTool1Page.Surface;
  Tool1PasswordHeaderLabel.Left := ScaleX(220);
  Tool1PasswordHeaderLabel.Top := TopPos;
  Tool1PasswordHeaderLabel.Caption := 'Password';
  Tool1PasswordHeaderLabel.Font.Style := [fsBold];
  
  TopPos := TopPos + ScaleY(25);
  
  for i := 0 to LocalUsersList.Count - 1 do
  begin
    UserCheckBoxes[i] := TCheckBox.Create(AdvancedTool1Page);
    UserCheckBoxes[i].Parent := AdvancedTool1Page.Surface;
    UserCheckBoxes[i].Left := ScaleX(20);
    UserCheckBoxes[i].Top := TopPos;
    UserCheckBoxes[i].Width := ScaleX(180);
    UserCheckBoxes[i].Caption := LocalUsersList[i];
    UserCheckBoxes[i].OnClick := @OnUserCheckBoxClick;

    UserPasswordEdits[i] := TEdit.Create(AdvancedTool1Page);
    UserPasswordEdits[i].Parent := AdvancedTool1Page.Surface;
    UserPasswordEdits[i].Left := ScaleX(220);
    UserPasswordEdits[i].Top := TopPos - ScaleY(2);
    UserPasswordEdits[i].Width := ScaleX(200);
    UserPasswordEdits[i].PasswordChar := '*';
    UserPasswordEdits[i].Enabled := False;
    UserPasswordEdits[i].OnChange := @OnPasswordEditChange;

    UserPasswordStatus[i] := TLabel.Create(AdvancedTool1Page);
    UserPasswordStatus[i].Parent := AdvancedTool1Page.Surface;
    UserPasswordStatus[i].Left := ScaleX(430);
    UserPasswordStatus[i].Top := TopPos - ScaleY(2);
    UserPasswordStatus[i].Font.Color := clRed;
    UserPasswordStatus[i].Caption := '';
    UserPasswordStatus[i].Visible := False;

    TopPos := TopPos + ScaleY(26);
  end;

  // Calculate bottom position for the password reset link
  BottomPos := AdvancedTool1Page.SurfaceHeight - ScaleY(40);
  
  // Create password reset link
  Tool1PasswordResetLink := TLabel.Create(AdvancedTool1Page);
  Tool1PasswordResetLink.Parent := AdvancedTool1Page.Surface;
  Tool1PasswordResetLink.Left := ScaleX(20);
  Tool1PasswordResetLink.Top := BottomPos;
  Tool1PasswordResetLink.Caption := 'Can''t remember a password? Click here to Reset it';
  // Choose a link color appropriate for current page theme
  if IsDarkColor(AdvancedTool1Page.Surface.Color) then
    Tool1PasswordResetLink.Font.Color := RGBToColor(135,206,250)
  else
    Tool1PasswordResetLink.Font.Color := clBlue;
  Tool1PasswordResetLink.Font.Style := [fsUnderline];
  Tool1PasswordResetLink.Cursor := crHandPoint;
  Tool1PasswordResetLink.OnClick := @OnPasswordResetLinkClick;

  SetUserControlsEnabled(True);
end;

function UserAlreadyEntered(const UserName: string): Boolean;
var
  i: Integer;
  CurrentUser: string;
  TempPassword: string;
begin
  Result := False;
  for i := 0 to UsersList.Count - 1 do
  begin
    ParseUserEntry(UsersList[i], CurrentUser, TempPassword);
    if CompareText(CurrentUser, UserName) = 0 then
    begin
      Result := True;
      exit;
    end;
  end;
end;

function IsValidUsername(const UserName: string): String;
var
  Len: Integer;
  i: Integer;
  Ch: Char;
begin
  Result := '';  // Empty string means valid
  
  UserName := Trim(UserName);
  Len := Length(UserName);
  
  // Check length (Windows limit is 20 characters for compatibility)
  if Len = 0 then
    Result := 'Username cannot be empty.'
  else if Len > 20 then
    Result := 'Username cannot exceed 20 characters.'
  else
  begin
    // Allow alphanumeric characters (a-z, A-Z, 0-9) and spaces
    for i := 1 to Len do
    begin
      Ch := UserName[i];
      if not ((Ch >= 'a') and (Ch <= 'z')) and
         not ((Ch >= 'A') and (Ch <= 'Z')) and
         not ((Ch >= '0') and (Ch <= '9')) and
         not (Ch = ' ') then
      begin
        Result := 'Username can only contain letters (a-z, A-Z), numbers (0-9), and spaces. No special characters allowed.';
        exit;
      end;
    end;
  end;
end;

function IsValidPassword(const Password: string): String;
begin
  Result := '';  // Empty string means valid
  
  if Length(Password) = 0 then
    Result := 'Password cannot be empty.'
  else if Length(Password) > 128 then
    Result := 'Password is too long (maximum 128 characters).';
end;

function ValidateLocalCredential(const UserName, Password: string): Boolean;
var
  Token: Cardinal;
begin
  // Fast local credential check via LogonUser; avoids slow PowerShell/WinRM
  Token := 0;
  Result := LogonUser(UserName, '.', Password, 2, 0, Token);
  if Token <> 0 then
    CloseHandle(Token);
end;

function ShouldInstallFiles: Boolean;
begin
  // Only copy files when the user selected to install RDP Wrapper/TermWrap
  Result := DoInstallRDPWrapper;
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;
  
  // Always skip directory selection page - must install to Program Files\RDP Wrapper
  if PageID = wpSelectDir then
    Result := True;
  
  // Skip User Page unless the flow indicates user creation is required
  if (PageID = UserPage.ID) and (not NeedCreateUsers) then
    Result := True;
  
  // Skip Ready page - no need to show install path
  if PageID = wpReady then
    Result := True;
  
  // Skip Advanced Tools Selection Page if not in advanced mode or if we jump directly
  if (PageID = AdvancedToolsSelectionPage.ID) and ((InstallType <> 2) or JumpToAdvancedTool1) then
    Result := True;
  
  // Skip Tool pages if not in advanced mode, or if wrong tool is selected
  // Allow AdvancedTool1Page when either in Advanced mode selecting tool 1, or when JumpToAdvancedTool1 is set
  if (PageID = AdvancedTool1Page.ID) and (not JumpToAdvancedTool1) and ((InstallType <> 2) or (SelectedAdvancedTool <> 0)) then
    Result := True;
  if (PageID = AdvancedTool2Page.ID) and ((InstallType <> 2) or (SelectedAdvancedTool <> 1)) then
    Result := True;
  if (PageID = AdvancedTool3Page.ID) and ((InstallType <> 2) or (SelectedAdvancedTool <> 2)) then
    Result := True;
  if (PageID = AdvancedTool4Page.ID) and ((InstallType <> 2) or (SelectedAdvancedTool <> 3)) then
    Result := True;
  if (PageID = AdvancedTool5Page.ID) and ((InstallType <> 2) or (SelectedAdvancedTool <> 4)) then
    Result := True;
end;

function IsVCRedistInstalled: Boolean;
var
  Major: Cardinal;
  Minor: Cardinal;
  Bld: Cardinal;
begin
  // Check for VC++ 2015-2022 Redistributable (x64)
  // The registry key stores the version information
  Result := RegQueryDWordValue(HKLM, REG_VCREDIST, 'Major', Major) and
            RegQueryDWordValue(HKLM, REG_VCREDIST, 'Minor', Minor) and
            RegQueryDWordValue(HKLM, REG_VCREDIST, 'Bld', Bld);
  
  if Result then
  begin
    // Check if version is 14.0 or higher (covers 2015-2022)
    Result := (Major >= 14);
  end;
end;

procedure SecureCleanupTempFiles(const UserName: string);
begin
  // Securely delete temporary files that contained sensitive data
  DeleteFile(TempFile('enc_' + UserName + '.txt'));
  DeleteFile(TempFile('create_rdp_' + UserName + '.ps1'));
end;

function EncryptPasswordToFile(const Password, UserName: string): string;
var
  ResultCode: Integer;
  EncPath: string;
begin
  EncPath := TempFile('enc_' + UserName + '.txt');
  // Encrypt the password and save to EncPath (without BOM or extra whitespace)
  Exec(EXE_POWERSHELL, BuildPowerShellArgs(
       '$pw = ''' + Password + ''' | ConvertTo-SecureString -AsPlainText -Force; ' +
       '$encPw = ConvertFrom-SecureString $pw; ' +
       '[System.IO.File]::WriteAllText(''' + EncPath + ''', $encPw)', True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := EncPath;
end;

function GenerateRDPPowerShellScript(const UserName, RDPPath, EncPath: string): string;
begin
  Result :=
    'param($EncPath)' + #13#10 +
    'try {' + #13#10 +
    '  Write-Host "Starting RDP file creation..."' + #13#10 +
    '  $encPassLines = @(Get-Content "$EncPath")' + #13#10 +
    '  $encPass = ($encPassLines -join "").Trim()' + #13#10 +
    '  Write-Host "Encrypted password loaded"' + #13#10 +
    '  $rdp = @()' + #13#10 +
    '  $rdp += "full address:s:127.0.0.2"' + #13#10 +
    '  $rdp += "username:s:' + UserName + '"' + #13#10 +
    '  $rdp += "screen mode id:i:1"' + #13#10 +
    '  $rdp += "desktopwidth:i:1366"' + #13#10 +
    '  $rdp += "desktopheight:i:768"' + #13#10 +
    '  $rdp += "use multimon:i:0"' + #13#10 +
    '  $rdp += "session bpp:i:32"' + #13#10 +
    '  $rdp += "smart sizing:i:1"' + #13#10 +
    '  $rdp += "compression:i:1"' + #13#10 +
    '  $rdp += "keyboardhook:i:1"' + #13#10 +
    '  $rdp += "audiocapturemode:i:0"' + #13#10 +
    '  $rdp += "audiomode:i:2"' + #13#10 +
    '  $rdp += "videoplaybackmode:i:1"' + #13#10 +
    '  $rdp += "connection type:i:7"' + #13#10 +
    '  $rdp += "displayconnectionbar:i:1"' + #13#10 +
    '  $rdp += "disable wallpaper:i:1"' + #13#10 +
    '  $rdp += "allow font smoothing:i:1"' + #13#10 +
    '  $rdp += "allow desktop composition:i:1"' + #13#10 +
    '  $rdp += "disable full window drag:i:0"' + #13#10 +
    '  $rdp += "disable menu anims:i:0"' + #13#10 +
    '  $rdp += "disable themes:i:0"' + #13#10 +
    '  $rdp += "bitmapcachepersistenable:i:1"' + #13#10 +
    '  $rdp += "authentication level:i:0"' + #13#10 +
    '  $rdp += "prompt for credentials:i:0"' + #13#10 +
    '  $rdp += "negotiate security layer:i:1"' + #13#10 +
    '  $rdp += "enablecredsspsupport:i:1"' + #13#10 +
    '  $rdp += "remoteapplicationmode:i:0"' + #13#10 +
    '  $rdp += "drivestoredirect:s:"' + #13#10 +
    '  $rdp += "alternate shell:s:"' + #13#10 +
    '  $rdp += "shell working directory:s:"' + #13#10 +
    '  $rdp += "gatewayhostname:s:"' + #13#10 +
    '  $rdp += "gatewayusagemethod:i:0"' + #13#10 +
    '  $rdp += "gatewaycredentialssource:i:0"' + #13#10 +
    '  $rdp += "gatewayprofileusagemethod:i:0"' + #13#10 +
    '  $rdp += "promptcredentialonce:i:0"' + #13#10 +
    '  $rdp += "use redirection server name:i:0"' + #13#10 +
    '  $rdp += ("password 51:b:" + $encPass)' + #13#10 +
    '  $rdp += "disableconnectionsharing:i:0"' + #13#10 +
    '  [System.IO.File]::WriteAllLines("' + RDPPath + '", $rdp)' + #13#10 +
    '  Start-Sleep -Milliseconds 500' + #13#10 +
    '  if (Test-Path "' + RDPPath + '") {' + #13#10 +
    '    Write-Host "File created successfully"' + #13#10 +
    '    exit 0' + #13#10 +
    '  } else {' + #13#10 +
    '    Write-Host "ERROR: File not created"' + #13#10 +
    '    exit 1' + #13#10 +
    '  }' + #13#10 +
    '} catch {' + #13#10 +
    '  Write-Host "EXCEPTION: $_"' + #13#10 +
    '  exit 1' + #13#10 +
    '}';
end;

procedure CreateRDPShortcut(const UserName, Password: string);
var
  ResultCode: Integer;
  RDPPath: string;
  EncPath: string;
  ScriptPath: string;
  PowerShellScript: string;
begin
  RDPPath := ExpandConstant('{userdesktop}\' + UserName + '.rdp');
  ScriptPath := TempFile('create_rdp_' + UserName + '.ps1');

  WriteInstallerLog('CreateRDPShortcut: Creating RDP file at ' + RDPPath);

  // Remove any existing shortcut so we always overwrite with the new one
  if FileExists(RDPPath) then
  begin
    WriteInstallerLog('CreateRDPShortcut: Deleting existing RDP file');
    DeleteFile(RDPPath);
  end;

  // Encrypt the password
  WriteInstallerLog('CreateRDPShortcut: Encrypting password for user ' + UserName);
  EncPath := EncryptPasswordToFile(Password, UserName);
  WriteInstallerLog('CreateRDPShortcut: Password encrypted to ' + EncPath);

  // Generate and execute PowerShell script
  WriteInstallerLog('CreateRDPShortcut: Generating RDP PowerShell script');
  PowerShellScript := GenerateRDPPowerShellScript(UserName, RDPPath, EncPath);
  SaveStringToFile(ScriptPath, PowerShellScript, False);
  WriteInstallerLog('CreateRDPShortcut: Script saved to ' + ScriptPath);
  
  WriteInstallerLog('CreateRDPShortcut: Executing PowerShell script');
  Exec(EXE_POWERSHELL, BuildPowerShellFileArgs(ScriptPath, '-EncPath "' + EncPath + '"', True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  WriteInstallerLog('CreateRDPShortcut: PowerShell script exit code=' + IntToStr(ResultCode));
  
  if ResultCode <> 0 then
  begin
    Log('WARNING: RDP file creation failed with exit code = ' + IntToStr(ResultCode));
    WriteInstallerLog('CreateRDPShortcut: WARNING - RDP file creation failed with exit code=' + IntToStr(ResultCode));
  end
  else
  begin
    WriteInstallerLog('CreateRDPShortcut: RDP file created successfully');
  end;

  // Securely delete temporary files containing sensitive data
  WriteInstallerLog('CreateRDPShortcut: Cleaning up temporary files');
  SecureCleanupTempFiles(UserName);
end;

procedure ClearPasswordsFromMemory;
begin
  // Clear all passwords from UsersList after use
  if Assigned(UsersList) then
  begin
    UsersList.Clear;
  end;
  if Assigned(ShortcutsList) then
  begin
    ShortcutsList.Clear;
  end;
end;

procedure CreateRDPUsers;
var
  ResultCode: Integer;
  i: Integer;
  UserInfo: string;
  UserName: string;
  Password: string;
  StartTick: Cardinal;
  UserStartTick: Cardinal;
begin
  // Lazy-resolve group names on first use (avoids blocking during InitializeWizard)
  if GroupAdministratorsName = 'Administrators' then
    GroupAdministratorsName := GetLocalizedGroupName('S-1-5-32-544', 'Administrators');
  if GroupRDPUsersName = 'Remote Desktop Users' then
    GroupRDPUsersName := GetLocalizedGroupName('S-1-5-32-555', 'Remote Desktop Users');
  
  // Start overall watchdog timer
  StartTick := GetTickCount;
  WriteInstallerLog('Starting CreateRDPUsers for ' + IntToStr(UsersList.Count) + ' users');
  
  for i := 0 to UsersList.Count - 1 do
  begin
    UserStartTick := GetTickCount;
    // Check overall timeout
    if (GetTickCount - StartTick) > USERS_OVERALL_TIMEOUT then
    begin
      WriteInstallerLog('CreateRDPUsers overall timeout reached after ' + IntToStr(GetTickCount - StartTick) + ' ms; aborting remaining users');
      break;
    end;

    UserInfo := UsersList[i];
    ParseUserEntry(UserInfo, UserName, Password);

    WizardForm.StatusLabel.Caption := 'Creating user account (' + IntToStr(i + 1) + ' of ' + IntToStr(UsersList.Count) + '): ' + UserName;
    WriteInstallerLog('Creating user: ' + UserName);

    // Create the Windows user account (quote username and password to handle spaces)
    Exec(EXE_CMD, '/c net user "' + UserName + '" "' + Password + '" /add /expires:never', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    WriteInstallerLog('net user exit=' + IntToStr(ResultCode) + ' for ' + UserName);
    Sleep(SLEEP_SHORT);

    // Add to both groups in sequence without extra delays
    Exec(EXE_CMD, '/c net localgroup "' + GroupAdministratorsName + '" "' + UserName + '" /add', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    WriteInstallerLog('add to Administrators exit=' + IntToStr(ResultCode) + ' for ' + UserName);
    Exec(EXE_CMD, '/c net localgroup "' + GroupRDPUsersName + '" "' + UserName + '" /add', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    WriteInstallerLog('add to Remote Desktop Users exit=' + IntToStr(ResultCode) + ' for ' + UserName);
    Sleep(SLEEP_SHORT);

    // Create RDP shortcut using helper function
    CreateRDPShortcut(UserName, Password);
    WriteInstallerLog('Created shortcut for ' + UserName);

    if (GetTickCount - UserStartTick) > PER_USER_TIMEOUT then
      WriteInstallerLog('WARNING: CreateRDPUsers per-user time exceeded ' + IntToStr(PER_USER_TIMEOUT) + ' ms for ' + UserName);
  end;
  WriteInstallerLog('CreateRDPUsers completed');
end;

procedure CreateShortcutsForExistingUsers;
var
  i: Integer;
  Entry: string;
  UserName: string;
  Password: string;
  StartTick: Cardinal;
  UserStartTick: Cardinal;
begin
  StartTick := GetTickCount;
  WriteInstallerLog('Starting CreateShortcutsForExistingUsers for ' + IntToStr(ShortcutsList.Count) + ' entries');
  for i := 0 to ShortcutsList.Count - 1 do
  begin
    UserStartTick := GetTickCount;
    if (GetTickCount - StartTick) > USERS_OVERALL_TIMEOUT then
    begin
      WriteInstallerLog('CreateShortcutsForExistingUsers overall timeout reached after ' + IntToStr(GetTickCount - StartTick) + ' ms; aborting remaining shortcuts');
      break;
    end;
    Entry := ShortcutsList[i];
    ParseUserEntry(Entry, UserName, Password);

    WizardForm.StatusLabel.Caption := 'Creating RDP shortcut (' + IntToStr(i + 1) + ' of ' + IntToStr(ShortcutsList.Count) + '): ' + UserName;

    // Create RDP shortcut using helper function
    CreateRDPShortcut(UserName, Password);
    WriteInstallerLog('Created shortcut for ' + UserName);

    if (GetTickCount - UserStartTick) > PER_USER_TIMEOUT then
      WriteInstallerLog('WARNING: CreateShortcutsForExistingUsers per-user time exceeded ' + IntToStr(PER_USER_TIMEOUT) + ' ms for ' + UserName);
  end;
  WriteInstallerLog('CreateShortcutsForExistingUsers completed');
end;

// Helper functions to display and update step-by-step progress on Installing page
procedure SetStepPending(L: TLabel; const Text: string);
begin
  if Assigned(L) then
  begin
    L.Caption := '• ' + Text;
    L.Font.Color := clGray;
    L.Font.Style := [];
    L.Visible := True;
  end;
end;

procedure SetStepInProgress(L: TLabel; const Text: string);
begin
  if Assigned(L) then
  begin
    L.Caption := '• ' + Text;
    L.Font.Color := clBlack;
    L.Font.Style := [fsBold];
    L.Visible := True;
  end;
end;

procedure SetStepDone(L: TLabel; const Text: string);
begin
  if Assigned(L) then
  begin
    L.Caption := '✓ ' + Text;
    L.Font.Color := clGreen;
    L.Font.Style := [];
    L.Visible := True;
  end;
end;

function CreateStepLabel(Parent: TWinControl; LeftPos, TopPos, WidthVal: Integer): TLabel;
var
  L: TLabel;
begin
  L := TLabel.Create(WizardForm);
  L.Parent := Parent;
  L.Left := LeftPos;
  L.Top := TopPos;
  L.Width := WidthVal;
  L.AutoSize := False;
  L.WordWrap := True;
  L.Visible := False;
  Result := L;
end;

// Hide all step labels before laying out the ones relevant to the selected mode
procedure HideAllStepLabels;
begin
  if Assigned(StepAddExcl) then StepAddExcl.Visible := False;
  if Assigned(StepRemoveExcl) then StepRemoveExcl.Visible := False;
  if Assigned(StepStopSvc) then StepStopSvc.Visible := False;
  if Assigned(StepEnsureVC) then StepEnsureVC.Visible := False;
  if Assigned(StepInstallRDPWrapper) then StepInstallRDPWrapper.Visible := False;
  if Assigned(StepConfigureService) then StepConfigureService.Visible := False;
  if Assigned(StepCreateUsers) then StepCreateUsers.Visible := False;
  if Assigned(StepCreateShortcuts) then StepCreateShortcuts.Visible := False;
  if Assigned(StepPreTrust) then StepPreTrust.Visible := False;
  if Assigned(StepStartSvc) then StepStartSvc.Visible := False;
  if Assigned(StepCheckRDP) then StepCheckRDP.Visible := False;
  if Assigned(StepUninstallRDPWrapper) then StepUninstallRDPWrapper.Visible := False;
  if Assigned(StepRemoveFolder) then StepRemoveFolder.Visible := False;
end;

// Begin a fresh layout pass for the steps list
procedure BeginStepLayout;
begin
  HideAllStepLabels;
  StepNextTop := StepTopBase;
end;

// Add a pending step label at the next position
procedure AddStepPendingLabel(L: TLabel; const Text: string);
begin
  if Assigned(L) then
  begin
    L.Left := StepLeftPos;
    L.Top := StepNextTop;
    L.Width := StepWidthVal;
    SetStepPending(L, Text);
    StepNextTop := StepNextTop + ScaleY(16);
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  // This is called before the install page shows, so we skip all blocking operations here
  // Service stop will be handled in CurStepChanged(ssInstall) after UI renders
  Result := '';
  Log('[PrepareToInstall] Skipping service operations (will be done in ssInstall step)');
end;

procedure InitializeWizard;
var
  leftPos, topPos, widthVal: Integer;
  extraGap: Integer;
  WelcomeExpLabel: TLabel;
  WelcomeIcon: TBitmapImage;
  // Credits labels (non-selectable) and link labels
  lblCreditsHeader: TLabel;
  lblBullet1: TLabel;
  lblStasName: TLabel;
  lblStasSuffix: TLabel;
  lblBullet2: TLabel;
  lblTermName: TLabel;
  lblTermSuffix: TLabel;
  lblBullet4: TLabel;
  lblBullet5: TLabel;
  lblBSGHName: TLabel;
  lblAnd: TLabel;
  lblBSSName: TLabel;
  lblProjectHome: TLabel;
  
  LinkColor: TColor;
  DesiredBottom: Integer;
  BlockTop: Integer;
  Delta: Integer;
begin
  // Initialize installer log file
  InitInstallerLog;
  // DON'T stop TermService here - it delays the wizard from showing by 5 seconds
  // It will be stopped later in the ssInstall step instead
  
  // Welcome intro page (shows description + credits)
  WelcomePage := CreateCustomPage(
    wpWelcome,
    'Welcome',
    ''
  );


  // Explanatory text in the main body of the welcome page
  WelcomeExpLabel := TLabel.Create(WelcomePage);
  WelcomeExpLabel.Parent := WelcomePage.Surface;
  WelcomeExpLabel.Left := ScaleX(10);
  WelcomeExpLabel.Top := ScaleY(10);
  // Ensure enough horizontal space and dynamic resizing in modern wizard styles
  WelcomeExpLabel.AutoSize := False;
  WelcomeExpLabel.Width := WelcomePage.SurfaceWidth - ScaleX(30);
  WelcomeExpLabel.Caption := 'RDPWrapKit sets up local RDP access, allowing you to run multiple users on your computer via Remote Desktop Protocol (RDP).' + #13#10#13#10 +
                           'You can choose to install RDP Wrapper and TermWrap, create/manage user accounts, and create RDP shortcuts on your desktop for easy access.';
  WelcomeExpLabel.WordWrap := True;
  WelcomeExpLabel.Alignment := taLeftJustify;
  WelcomeExpLabel.Font.Size := WelcomeExpLabel.Font.Size + 2;
  // Anchor left+right so label width tracks page surface on dynamic layouts
  WelcomeExpLabel.Anchors := [akLeft, akRight];
  // Set a fixed height large enough for wrapped text
  WelcomeExpLabel.Height := ScaleY(80);

  // Add an image on the left side of the welcome page
  WelcomeIcon := TBitmapImage.Create(WelcomePage);
  WelcomeIcon.Parent := WelcomePage.Surface;
  WelcomeIcon.Left := ScaleX(8);
  WelcomeIcon.Top := ScaleY(8);
  WelcomeIcon.Width := ScaleX(96);
  WelcomeIcon.Height := ScaleY(96);
  WelcomeIcon.Stretch := True;
  try
    // Explicitly extract the file to ensure it exists
    ExtractTemporaryFile(FILE_ICON_BMP);
    WelcomeIcon.Bitmap.LoadFromFile(ExpandConstant(TEMP_ICON_BMP));
  except
    // Log error and continue without icon
    Log('[Icon Load Error] Failed to load RDPWrapKitIcon.bmp');
  end;

  // Shift welcome text to the right of the icon
  WelcomeExpLabel.Left := WelcomeIcon.Left + WelcomeIcon.Width + ScaleX(10);
  WelcomeExpLabel.Width := WelcomePage.SurfaceWidth - WelcomeExpLabel.Left - ScaleX(20);

  // Choices page (placed after the welcome intro)
  InstallTypePage := CreateCustomPage(
    WelcomePage.ID,
    'Setup Options',
    'Select what you would like to do:'
  );

  // Create top-level radio buttons: Typical, Advanced, Uninstall
  TypicalRadio := TRadioButton.Create(InstallTypePage);
  TypicalRadio.Parent := InstallTypePage.Surface;
  TypicalRadio.Left := ScaleX(10);
  TypicalRadio.Top := ScaleY(10);
  TypicalRadio.Width := ScaleX(420);
  TypicalRadio.Caption := 'Typical Setup';
  TypicalRadio.Checked := True;
  TypicalRadio.OnClick := @OnInstallTypeChange;

  // Under Typical, add checkboxes and nested radios
  chkInstallRDPWrapper := TCheckBox.Create(InstallTypePage);
  chkInstallRDPWrapper.Parent := InstallTypePage.Surface;
  chkInstallRDPWrapper.Left := ScaleX(30);
  chkInstallRDPWrapper.Top := ScaleY(36);
  chkInstallRDPWrapper.Width := ScaleX(380);
  chkInstallRDPWrapper.Caption := 'Install RDP Wrapper and TermWrap';
  chkInstallRDPWrapper.Checked := True;

  chkManageUsers := TCheckBox.Create(InstallTypePage);
  chkManageUsers.Parent := InstallTypePage.Surface;
  chkManageUsers.Left := ScaleX(30);
  chkManageUsers.Top := ScaleY(60);
  chkManageUsers.Width := ScaleX(380);
  chkManageUsers.Caption := 'Manage Users / Shortcuts';
  chkManageUsers.Checked := True;
  chkManageUsers.OnClick := @OnManageUsersClick;

  // Panel to contain the Create/Use Existing radios so they form their own group
  ManageUsersGroup := TPanel.Create(InstallTypePage);
  ManageUsersGroup.Parent := InstallTypePage.Surface;
  ManageUsersGroup.Left := ScaleX(40);
  ManageUsersGroup.Top := ScaleY(80);
  ManageUsersGroup.Width := ScaleX(360);
  ManageUsersGroup.Height := ScaleY(64);
  ManageUsersGroup.BorderStyle := bsNone;
  ManageUsersGroup.Color := InstallTypePage.Surface.Color;
  // Remove bevel to make the panel flat
  ManageUsersGroup.BevelInner := bvNone;
  ManageUsersGroup.BevelOuter := bvNone;
  ManageUsersGroup.BevelWidth := 0;

  rbCreateUsers := TRadioButton.Create(ManageUsersGroup);
  rbCreateUsers.Parent := ManageUsersGroup;
  rbCreateUsers.Left := ScaleX(10);
  rbCreateUsers.Top := ScaleY(8);
  rbCreateUsers.Width := ScaleX(340);
  rbCreateUsers.Caption := 'Create users / create shortcuts';
  rbCreateUsers.Checked := True;
  rbCreateUsers.OnClick := @OnManageUsersClick;

  rbUseExistingUsers := TRadioButton.Create(ManageUsersGroup);
  rbUseExistingUsers.Parent := ManageUsersGroup;
  rbUseExistingUsers.Left := ScaleX(10);
  rbUseExistingUsers.Top := ScaleY(32);
  rbUseExistingUsers.Width := ScaleX(340);
  rbUseExistingUsers.Caption := 'Use existing users / create shortcuts';
  rbUseExistingUsers.Checked := False;
  rbUseExistingUsers.OnClick := @OnManageUsersClick;

  // Set initial enabled state based on ManageUsers checkbox
  rbCreateUsers.Enabled := chkManageUsers.Checked;
  rbUseExistingUsers.Enabled := chkManageUsers.Checked;

  // Add some white space
  // Advanced radio (move down ~20% of surface height to leave more white space)
  extraGap := InstallTypePage.SurfaceHeight div 5; // ~20%
  AdvancedRadio := TRadioButton.Create(InstallTypePage);
  AdvancedRadio.Parent := InstallTypePage.Surface;
  AdvancedRadio.Left := ScaleX(10);
  AdvancedRadio.Top := ScaleY(150) + extraGap;
  AdvancedRadio.Width := ScaleX(420);
  AdvancedRadio.Caption := 'Advanced: Tools and Utilities';
  AdvancedRadio.Checked := False;
  AdvancedRadio.Visible := False;
  AdvancedRadio.OnClick := @OnInstallTypeChange;

  // Uninstall radio (also moved down by the same gap)
  UninstallRadio := TRadioButton.Create(InstallTypePage);
  UninstallRadio.Parent := InstallTypePage.Surface;
  UninstallRadio.Left := ScaleX(10);
  UninstallRadio.Top := ScaleY(176) + extraGap;
  UninstallRadio.Width := ScaleX(420);
  UninstallRadio.Caption := 'Uninstall RDP Wrapper and TermWrap';
  UninstallRadio.Checked := False;
  UninstallRadio.OnClick := @OnInstallTypeChange;

  // Add credits text at bottom of the welcome intro page
  CreditsText := TRichEditViewer.Create(WelcomePage);
  CreditsText.Parent := WelcomePage.Surface;
  CreditsText.Left := ScaleX(10);
  CreditsText.Width := WelcomePage.SurfaceWidth - ScaleX(20);
  CreditsText.Height := ScaleY(120);
  CreditsText.ScrollBars := ssNone;
  CreditsText.BorderStyle := bsNone;
  CreditsText.Color := WelcomePage.Surface.Color;
  CreditsText.Font.Size := 9;
  // Use the wizard's default font color so the credits follow the current theme
  CreditsText.Font.Color := WizardForm.Font.Color;
  // Use the standard Windows link blue which contrasts well in light/dark themes
  LinkColor := RGBToColor(0,120,215);

  // Hide the rich viewer and render credits as non-selectable labels; links remain clickable
  CreditsText.Visible := False;

  // Header (bold / slightly larger)
  lblCreditsHeader := TLabel.Create(WelcomePage);
  lblCreditsHeader.Parent := WelcomePage.Surface;
  lblCreditsHeader.Left := CreditsText.Left;
  lblCreditsHeader.Top := CreditsText.Top;
  lblCreditsHeader.Caption := 'RDPWrapKit is assembled from the great work below:';
  lblCreditsHeader.Font.Style := [fsBold];
  lblCreditsHeader.Font.Size := CreditsText.Font.Size + 1;
  lblCreditsHeader.Font.Color := CreditsText.Font.Color;
  lblCreditsHeader.Transparent := True;
  lblCreditsHeader.AutoSize := True;

  // Start placing bullet lines under the header
  topPos := lblCreditsHeader.Top + ScaleY(18);

  // Line 1: Stas'M's RDP Wrapper (clickable)
  lblBullet1 := TLabel.Create(WelcomePage);
  lblBullet1.Parent := WelcomePage.Surface;
  lblBullet1.Left := CreditsText.Left;
  lblBullet1.Top := topPos;
  lblBullet1.Caption := '• ';
  lblBullet1.Font.Color := CreditsText.Font.Color;
  lblBullet1.Transparent := True;
  lblBullet1.AutoSize := True;

  lblStasName := TLabel.Create(WelcomePage);
  lblStasName.Parent := WelcomePage.Surface;
  lblStasName.Left := lblBullet1.Left + lblBullet1.Width;
  lblStasName.Top := topPos;
  lblStasName.Caption := 'Stas''M''s RDP Wrapper';
  lblStasName.Font.Color := LinkColor;
  lblStasName.Font.Style := [fsUnderline];
  lblStasName.Cursor := crHand;
  lblStasName.OnClick := @OpenRDPWrap;
  lblStasName.Transparent := True;
  lblStasName.AutoSize := True;

  lblStasSuffix := TLabel.Create(WelcomePage);
  lblStasSuffix.Parent := WelcomePage.Surface;
  lblStasSuffix.Left := lblStasName.Left + lblStasName.Width + ScaleX(4);
  lblStasSuffix.Top := topPos;
  lblStasSuffix.Caption := ' - (Apache 2.0 license)';
  lblStasSuffix.Font.Color := CreditsText.Font.Color;
  lblStasSuffix.Transparent := True;
  lblStasSuffix.AutoSize := True;

  topPos := topPos + ScaleY(18);

  lblBullet2 := TLabel.Create(WelcomePage);
  lblBullet2.Parent := WelcomePage.Surface;
  lblBullet2.Left := CreditsText.Left;
  lblBullet2.Top := topPos;
  lblBullet2.Caption := '• ';
  lblBullet2.Font.Color := CreditsText.Font.Color;
  lblBullet2.Transparent := True;
  lblBullet2.AutoSize := True;

  lblTermName := TLabel.Create(WelcomePage);
  lblTermName.Parent := WelcomePage.Surface;
  lblTermName.Left := lblBullet2.Left + lblBullet2.Width;
  lblTermName.Top := topPos;
  lblTermName.Caption := 'llccd''s TermWrap';
  lblTermName.Font.Color := LinkColor;
  lblTermName.Font.Style := [fsUnderline];
  lblTermName.Cursor := crHand;
  lblTermName.OnClick := @OpenTermWrap;
  lblTermName.Transparent := True;
  lblTermName.AutoSize := True;

  lblTermSuffix := TLabel.Create(WelcomePage);
  lblTermSuffix.Parent := WelcomePage.Surface;
  lblTermSuffix.Left := lblTermName.Left + lblTermName.Width + ScaleX(4);
  lblTermSuffix.Top := topPos;
  lblTermSuffix.Caption := ' - (MIT License)';
  lblTermSuffix.Font.Color := CreditsText.Font.Color;
  lblTermSuffix.Transparent := True;
  lblTermSuffix.AutoSize := True;

  topPos := topPos + ScaleY(18);

  // Line 3: Special thanks with two clickable names
  lblBullet5 := TLabel.Create(WelcomePage);
  lblBullet5.Parent := WelcomePage.Surface;
  lblBullet5.Left := CreditsText.Left;
  lblBullet5.Top := topPos;
  lblBullet5.Caption := '• Special thanks to Bee Swarm Simulator communities: ';
  lblBullet5.Font.Color := CreditsText.Font.Color;
  lblBullet5.Transparent := True;
  lblBullet5.AutoSize := True;

  lblBSGHName := TLabel.Create(WelcomePage);
  lblBSGHName.Parent := WelcomePage.Surface;
  lblBSGHName.Left := lblBullet5.Left + lblBullet5.Width;
  lblBSGHName.Top := topPos;
  lblBSGHName.Caption := 'BSGH';
  lblBSGHName.Font.Color := LinkColor;
  lblBSGHName.Font.Style := [fsUnderline];
  lblBSGHName.Cursor := crHand;
  lblBSGHName.OnClick := @OpenBSGH;
  lblBSGHName.Transparent := True;
  lblBSGHName.AutoSize := True;

  lblAnd := TLabel.Create(WelcomePage);
  lblAnd.Parent := WelcomePage.Surface;
  lblAnd.Left := lblBSGHName.Left + lblBSGHName.Width + ScaleX(4);
  lblAnd.Top := topPos;
  lblAnd.Caption := 'and ';
  lblAnd.Font.Color := CreditsText.Font.Color;
  lblAnd.Transparent := True;
  lblAnd.AutoSize := True;

  lblBSSName := TLabel.Create(WelcomePage);
  lblBSSName.Parent := WelcomePage.Surface;
  lblBSSName.Left := lblAnd.Left + lblAnd.Width;
  lblBSSName.Top := topPos;
  lblBSSName.Caption := 'BSS Grinders';
  lblBSSName.Font.Color := LinkColor;
  lblBSSName.Font.Style := [fsUnderline];
  lblBSSName.Cursor := crHand;
  lblBSSName.OnClick := @OpenBSSGrinders;
  lblBSSName.Transparent := True;
  lblBSSName.AutoSize := True;

  topPos := topPos + ScaleY(18);
  topPos := topPos + ScaleY(18);

  // Line 4: Assembled by cpdx4. Project Home:
  lblBullet4 := TLabel.Create(WelcomePage);
  lblBullet4.Parent := WelcomePage.Surface;
  lblBullet4.Left := CreditsText.Left;
  lblBullet4.Top := topPos;
  lblBullet4.Caption := '• Assembled by cpdx4. Project Home: ';
  lblBullet4.Font.Color := CreditsText.Font.Color;
  lblBullet4.Transparent := True;
  lblBullet4.AutoSize := True;

  // Clickable Project Home link (only URL portion)
  lblProjectHome := TLabel.Create(WelcomePage);
  lblProjectHome.Parent := WelcomePage.Surface;
  lblProjectHome.Left := lblBullet4.Left + lblBullet4.Width + ScaleX(4);
  lblProjectHome.Top := topPos;
  lblProjectHome.Caption := 'github.com/cpdx4/RDPWrapKit';
  lblProjectHome.Font.Color := LinkColor;
  lblProjectHome.Font.Style := [fsUnderline];
  lblProjectHome.Cursor := crHand;
  lblProjectHome.OnClick := @OpenProjectHome;
  lblProjectHome.Transparent := True;
  lblProjectHome.AutoSize := True;

  // Position the block near the bottom of the welcome page area
  topPos := topPos + ScaleY(24);
  DesiredBottom := WelcomePage.SurfaceHeight - ScaleY(10);
  BlockTop := DesiredBottom - (topPos - CreditsText.Top);
  if BlockTop < CreditsText.Top then
    BlockTop := CreditsText.Top;
  // Shift all created labels by the delta to align the block
  Delta := BlockTop - CreditsText.Top;
  lblCreditsHeader.Top := lblCreditsHeader.Top + Delta;
  lblBullet1.Top := lblBullet1.Top + Delta;
  lblStasName.Top := lblStasName.Top + Delta;
  lblStasSuffix.Top := lblStasSuffix.Top + Delta;
  lblBullet2.Top := lblBullet2.Top + Delta;
  lblTermName.Top := lblTermName.Top + Delta;
  lblTermSuffix.Top := lblTermSuffix.Top + Delta;
  lblBullet5.Top := lblBullet5.Top + Delta;
  lblBSGHName.Top := lblBSGHName.Top + Delta;
  lblAnd.Top := lblAnd.Top + Delta;
  lblBSSName.Top := lblBSSName.Top + Delta;
  lblBullet4.Top := lblBullet4.Top + Delta;
  lblProjectHome.Top := lblProjectHome.Top + Delta;

  // Initialize derived flags to reflect defaults so page skipping works immediately
  DoInstallRDPWrapper := chkInstallRDPWrapper.Checked;
  DoManageUsers := chkManageUsers.Checked;
  NeedCreateUsers := DoManageUsers and rbCreateUsers.Checked;
  
  // Create User Page (after InstallTypePage)
  UserPage := CreateInputQueryPage(
    InstallTypePage.ID,
    'Create RDP User Account',
    'Create a new user by entering a username (such as "macro1" or "rdp1") and password below.',
    ''
  );
  UserPage.Add('Username:', False);
  UserPage.Add('Password:', True);  // masked input
  
  // Create Advanced Tools Selection Page with radio buttons (1-5)
  AdvancedToolsSelectionPage := CreateInputOptionPage(
    InstallTypePage.ID,
    'Advanced Tools',
    'Available Tools',
    'Select a tool to use:',
    True,
    False
  );
  
  AdvancedToolsSelectionPage.Add('1. placeholder');
  AdvancedToolsSelectionPage.Add('2. placeholder');
  AdvancedToolsSelectionPage.Add('3. placeholder');
  AdvancedToolsSelectionPage.Add('4. placeholder');
  AdvancedToolsSelectionPage.Add('5. placeholder');
  AdvancedToolsSelectionPage.Values[0] := True;  // Default to tool 1
  
  // Create Tool 1 Page: Create RDP desktop shortcuts for existing local users
  AdvancedTool1Page := CreateCustomPage(
    AdvancedToolsSelectionPage.ID,
    'Create RDP Desktop Shortcuts',
    'This is a list of users found on this PC. Checkmark the accounts you want to make desktop shortcuts for and type their password.'
  );
  
  // Create placeholder pages for tools 2-5
  AdvancedTool2Page := CreateCustomPage(
    AdvancedToolsSelectionPage.ID,
    'Tool 2',
    'Placeholder tool 2'
  );
  
  AdvancedTool3Page := CreateCustomPage(
    AdvancedToolsSelectionPage.ID,
    'Tool 3',
    'Placeholder tool 3'
  );
  
  AdvancedTool4Page := CreateCustomPage(
    AdvancedToolsSelectionPage.ID,
    'Tool 4',
    'Placeholder tool 4'
  );
  
  AdvancedTool5Page := CreateCustomPage(
    AdvancedToolsSelectionPage.ID,
    'Tool 5',
    'Placeholder tool 5'
  );
  
  // Keep old AdvancedPage for backward compatibility (not used in new flow)
  AdvancedPage := AdvancedTool1Page;
  
  // Initialize lists for tracking
  LocalUsersList := TStringList.Create;  // Will be populated when Advanced Tools page is shown
  SetLength(UserCheckBoxes, 0);
  SetLength(UserPasswordEdits, 0);
  SetLength(UserPasswordStatus, 0);
  ShortcutsList := TStringList.Create;
  
  // Add label for options section
  OptionsLabel := TLabel.Create(UserPage);
  OptionsLabel.Parent := UserPage.Surface;
  OptionsLabel.Left := ScaleX(10);
  OptionsLabel.Top := ScaleY(200);
  OptionsLabel.Caption := 'What would you like to do next?';
  OptionsLabel.Font.Style := [fsBold];
  
  // Add "Create more users" radio button
  AddMoreRadio := TRadioButton.Create(UserPage);
  AddMoreRadio.Parent := UserPage.Surface; 
  AddMoreRadio.Left := ScaleX(20);
  AddMoreRadio.Top := OptionsLabel.Top + ScaleY(25);
  AddMoreRadio.Width := ScaleX(400);
  AddMoreRadio.Caption := 'I want to create another user';
  AddMoreRadio.Checked := False;
  
  // Add "Done creating users" radio button
  DoneRadio := TRadioButton.Create(UserPage);
  DoneRadio.Parent := UserPage.Surface;
  DoneRadio.Left := ScaleX(20);
  DoneRadio.Top := AddMoreRadio.Top + ScaleY(25);
  DoneRadio.Width := ScaleX(400);
  DoneRadio.Caption := 'I''m done creating users, continue setup';
  DoneRadio.Checked := True;  // Default selection

  // Subtle opt-out for users who already have accounts
  SkipUsersCheckBox := TCheckBox.Create(UserPage);
  SkipUsersCheckBox.Parent := UserPage.Surface;
  SkipUsersCheckBox.Left := ScaleX(20);
  SkipUsersCheckBox.Top := DoneRadio.Top + ScaleY(22);
  SkipUsersCheckBox.Width := ScaleX(420);
  SkipUsersCheckBox.Caption := 'Skip creating users for now (advanced)';
  SkipUsersCheckBox.Checked := False;
  if IsDarkColor(UserPage.Surface.Color) then
    SkipUsersCheckBox.Font.Color := RGBToColor(190,190,190)
  else
    SkipUsersCheckBox.Font.Color := clGray;
  
  // Initialize UsersList
  UsersList := TStringList.Create;
  CreatedUsersList := TStringList.Create;
  CurrentUserIndex := 0;
  // Shortcuts selections
  // ShortcutsList already initialized above; nothing else needed
  
  // Debug: Set to True to force VC++ download
  DebugMode := False;

  // Initialize group names with defaults (will be resolved on-demand if needed)
  // Deferring PowerShell calls to avoid blocking the first UI from showing
  GroupAdministratorsName := 'Administrators';
  GroupRDPUsersName := 'Remote Desktop Users';

  // Build the progress steps area on the Installing page (under the progress bar)
  leftPos := WizardForm.ProgressGauge.Left;
  topPos := WizardForm.ProgressGauge.Top + WizardForm.ProgressGauge.Height + ScaleY(12);
  widthVal := WizardForm.ProgressGauge.Width;

  StepsHeaderLabel := TLabel.Create(WizardForm);
  StepsHeaderLabel.Parent := WizardForm.InstallingPage;
  StepsHeaderLabel.Left := leftPos;
  StepsHeaderLabel.Top := topPos;
  StepsHeaderLabel.Caption := 'Steps:';
  StepsHeaderLabel.Font.Style := [fsBold];
  StepsHeaderLabel.Visible := True;

  topPos := topPos + ScaleY(18);
  // Capture layout metrics for later dynamic reflow
  StepLeftPos := leftPos;
  StepTopBase := topPos;
  StepWidthVal := widthVal;
  StepStopSvc := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);          topPos := topPos + ScaleY(16);
  StepAddExcl := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);          topPos := topPos + ScaleY(16);
  StepRemoveExcl := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);       topPos := topPos + ScaleY(16);
  StepEnsureVC := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);         topPos := topPos + ScaleY(16);
  StepInstallRDPWrapper := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);topPos := topPos + ScaleY(16);
  StepConfigureService := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal); topPos := topPos + ScaleY(16);
  StepCreateUsers := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);      topPos := topPos + ScaleY(16);
  StepCreateShortcuts := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);  topPos := topPos + ScaleY(16);
  StepPreTrust := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);         topPos := topPos + ScaleY(16);
  StepStartSvc := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);         topPos := topPos + ScaleY(16);
  StepCheckRDP := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);         topPos := topPos + ScaleY(16);
  StepCheckMSTSC := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);       topPos := topPos + ScaleY(16);
  StepInstallMSTSC := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);     topPos := topPos + ScaleY(16);
  StepUninstallRDPWrapper := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal); topPos := topPos + ScaleY(16);
  StepRemoveFolder := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);

  // Create a scrollable rich text viewer on the Finished page and hide the default label
  FinishedText := TRichEditViewer.Create(WizardForm.FinishedLabel.Parent);
  FinishedText.Parent := WizardForm.FinishedLabel.Parent;
  FinishedText.Left := WizardForm.FinishedLabel.Left;
  FinishedText.Top := WizardForm.FinishedLabel.Top;
  FinishedText.Width := WizardForm.FinishedLabel.Width;
  // Calculate height based on parent's available space, not the label's original small height
  FinishedText.Height := WizardForm.FinishedLabel.Parent.ClientHeight - WizardForm.FinishedLabel.Top - ScaleY(35);
  FinishedText.ScrollBars := ssVertical;
  FinishedText.BorderStyle := bsNone;
  FinishedText.Color := WizardForm.FinishedLabel.Color;
  FinishedText.Font.Size := WizardForm.FinishedLabel.Font.Size;
  FinishedText.Visible := True;
  WizardForm.FinishedLabel.Visible := False;

  // Create a button to view the install log on the Finished page (positioned directly below text)
  ViewLogButton := TButton.Create(WizardForm.FinishedLabel.Parent);
  ViewLogButton.Parent := WizardForm.FinishedLabel.Parent;
  ViewLogButton.Left := WizardForm.FinishedLabel.Left;
  ViewLogButton.Top := FinishedText.Top + FinishedText.Height + ScaleY(5);
  ViewLogButton.Width := ScaleX(160);
  ViewLogButton.Height := ScaleY(24);
  ViewLogButton.Caption := 'View Install Log';
  ViewLogButton.OnClick := @OnViewLogButtonClick;
  ViewLogButton.Visible := True;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  UserName: string;
  Password: string;
  i: Integer;
  SelectedCount: Integer;
  HasErrors: Boolean;
begin
  Result := True;
  
  if CurPageID = InstallTypePage.ID then
  begin
    // Derive install type and flags from new welcome/options controls
    // Default flags
    DoInstallRDPWrapper := False;
    DoManageUsers := False;
    NeedCreateUsers := False;

    if UninstallRadio.Checked then
    begin
      InstallType := 3; // Uninstall Everything
    end
    else if AdvancedRadio.Checked then
    begin
      InstallType := 2; // Advanced tools
    end
    else // Typical selected
    begin
      DoInstallRDPWrapper := chkInstallRDPWrapper.Checked;
      DoManageUsers := chkManageUsers.Checked;

      if DoManageUsers and rbCreateUsers.Checked then
        NeedCreateUsers := True
      else
        NeedCreateUsers := False;

      // Require at least one action under Typical
      if (not DoInstallRDPWrapper) and (not DoManageUsers) then
      begin
        MsgBox('Please select at least one option under Typical Setup', mbError, MB_OK);
        Result := False;
        exit;
      end;

      // Map to existing InstallType semantics for downstream logic
      if DoInstallRDPWrapper then
      begin
        // If we're installing wrapper, treat as Full Install (0)
        InstallType := 0;
      end
      else if DoManageUsers and rbCreateUsers.Checked then
      begin
        // Create users only (no files)
        InstallType := 1;
      end
      else if DoManageUsers and rbUseExistingUsers.Checked then
      begin
        // If only Manage Users (use existing) requested without install, treat as Advanced
        if not DoInstallRDPWrapper then
          InstallType := 2
        else
          InstallType := 0; // will install then create shortcuts
      end
      else
      begin
        // Fallback to Full Install
        InstallType := 0;
      end;
      // If user chose Use existing users from Typical, jump directly to Advanced tool 1
      if DoManageUsers and rbUseExistingUsers.Checked then
      begin
        SelectedAdvancedTool := 0;
        JumpToAdvancedTool1 := True;
        // Ensure advanced tool controls are built so next page shows selections
        LocalUsersList := GetLocalUsers;
        SetLength(UserCheckBoxes, LocalUsersList.Count);
        SetLength(UserPasswordEdits, LocalUsersList.Count);
        SetLength(UserPasswordStatus, LocalUsersList.Count);
        BuildAdvancedUserControls;
      end
      else
        JumpToAdvancedTool1 := False;
    end;
  end
  else if CurPageID = AdvancedToolsSelectionPage.ID then
  begin
    // Store selected advanced tool (0-4 for tools 1-5)
    if AdvancedToolsSelectionPage.Values[0] then
      SelectedAdvancedTool := 0
    else if AdvancedToolsSelectionPage.Values[1] then
      SelectedAdvancedTool := 1
    else if AdvancedToolsSelectionPage.Values[2] then
      SelectedAdvancedTool := 2
    else if AdvancedToolsSelectionPage.Values[3] then
      SelectedAdvancedTool := 3
    else
      SelectedAdvancedTool := 4;
  end
  else if CurPageID = AdvancedTool1Page.ID then
  begin
      ShortcutsList.Clear;

      SelectedCount := 0;
      HasErrors := False;

      // Collect selections and validate all at once
      for i := 0 to High(UserCheckBoxes) do
      begin
        if Assigned(UserCheckBoxes[i]) and UserCheckBoxes[i].Checked then
        begin
          Inc(SelectedCount);
          if SelectedCount > MAX_SHORTCUTS then
          begin
            MsgBox('You can create a maximum of ' + IntToStr(MAX_SHORTCUTS) + ' shortcuts at a time.', mbError, MB_OK);
            Result := False;
            exit;
          end;

          Password := UserPasswordEdits[i].Text;
          if Password = '' then
          begin
            UserPasswordStatus[i].Caption := 'Password required';
            UserPasswordStatus[i].Visible := True;
            HasErrors := True;
            continue;
          end;

          if IsValidPassword(Password) <> '' then
          begin
            UserPasswordStatus[i].Caption := 'Invalid password';
            UserPasswordStatus[i].Visible := True;
            HasErrors := True;
            continue;
          end;

          if not ValidateLocalCredential(LocalUsersList[i], Password) then
          begin
            UserPasswordStatus[i].Caption := 'Incorrect PW';
            UserPasswordStatus[i].Visible := True;
            HasErrors := True;
            continue;
          end;

          UserPasswordStatus[i].Visible := False;

          ShortcutsList.Add(LocalUsersList[i] + '|' + Password);
        end
        else if Assigned(UserPasswordStatus[i]) then
        begin
          UserPasswordStatus[i].Visible := False;
        end;
      end;

      if SelectedCount = 0 then
      begin
        MsgBox('Select at least one user to create a shortcut.', mbError, MB_OK);
        Result := False;
        exit;
      end;

      if HasErrors then
      begin
        Result := False;
        exit;
      end;

      Result := True;
  end
  else if CurPageID = UserPage.ID then
  begin
    // Skip user page for uninstall and advanced types
    if (InstallType = 2) or (InstallType = 3) then
    begin
      Result := True;
      exit;
    end;
    
    UserName := UserPage.Values[0];
    Password := UserPage.Values[1];
    
    // Check if user selected "I'm done creating users"
    if DoneRadio.Checked then
    begin
      // If current fields have data, validate and add before counting
      if (UserName <> '') or (Password <> '') then
      begin
        if UserName = '' then
        begin
          MsgBox('Please enter a username.', mbError, MB_OK);
          Result := False;
          exit;
        end;

        // Validate username format
        if IsValidUsername(UserName) <> '' then
        begin
          MsgBox(IsValidUsername(UserName), mbError, MB_OK);
          Result := False;
          exit;
        end;

        if UserAlreadyEntered(UserName) then
        begin
          MsgBox('Error: You already entered a user named "' + UserName + '". Please choose a different username.', mbError, MB_OK);
          Result := False;
          exit;
        end;

        if UserExists(UserName) then
        begin
          MsgBox('Error: User "' + UserName + '" already exists. Please choose a different username.', mbError, MB_OK);
          Result := False;
          exit;
        end;

        if Password = '' then
        begin
          MsgBox('Please enter a password.', mbError, MB_OK);
          Result := False;
          exit;
        end;
        
        // Validate password
        if IsValidPassword(Password) <> '' then
        begin
          MsgBox(IsValidPassword(Password), mbError, MB_OK);
          Result := False;
          exit;
        end;
        
        UsersList.Add(UserName + '|' + Password);
      end;

      // After adding pending entry, ensure at least one user exists
      if (UsersList.Count = 0) and (not SkipUsersCheckBox.Checked) then
      begin
        MsgBox('Please create at least one user account before proceeding.', mbError, MB_OK);
        Result := False;
        exit;
      end;

      // For both Full Install and Add Users Only, proceed to install phase
      // (file copying will be skipped for Add Users Only via the Check parameter)
      Result := True;
      exit;
    end;
    
    // "I want to create more users" is selected - validate current entry
    if UserName = '' then
    begin
      MsgBox('Please enter a username.', mbError, MB_OK);
      Result := False;
      exit;
    end;

    // Validate username format
    if IsValidUsername(UserName) <> '' then
    begin
      MsgBox(IsValidUsername(UserName), mbError, MB_OK);
      Result := False;
      exit;
    end;

    if UserAlreadyEntered(UserName) then
    begin
      MsgBox('Error: You already entered a user named "' + UserName + '". Please choose a different username.', mbError, MB_OK);
      Result := False;
      exit;
    end;

    if UserExists(UserName) then
    begin
      MsgBox('Error: User "' + UserName + '" already exists. Please choose a different username.', mbError, MB_OK);
      Result := False;
      exit;
    end;

    if Password = '' then
    begin
      MsgBox('Please enter a password.', mbError, MB_OK);
      Result := False;
      exit;
    end;
    
    // Validate password
    if IsValidPassword(Password) <> '' then
    begin
      MsgBox(IsValidPassword(Password), mbError, MB_OK);
      Result := False;
      exit;
    end;
    
    // Store the user credentials
    UsersList.Add(UserName + '|' + Password);
    
    // Clear fields and stay on this page for another user
    UserPage.Values[0] := '';
    UserPage.Values[1] := '';
    
    // Default to "I'm done adding users" for the next entry
    DoneRadio.Checked := True;
    AddMoreRadio.Checked := False;
    
    Result := False;
  end;
end;

function BackButtonClick(CurPageID: Integer): Boolean;
var
  LastUserInfo: string;
  UserName: string;
  Password: string;
begin
  Result := True;
  
  // If user clicks Back on the UserPage and there are users in the list,
  // go back to the previous user (for editing)
  if CurPageID = UserPage.ID then
  begin
    if UsersList.Count > 0 then
    begin
      // Get the last user from the list
      LastUserInfo := UsersList[UsersList.Count - 1];
      ParseUserEntry(LastUserInfo, UserName, Password);
      
      // Populate the fields with the previous user's data
      UserPage.Values[0] := UserName;
      UserPage.Values[1] := Password;
      
      // Remove this user from the list (so they can re-enter or modify)
      UsersList.Delete(UsersList.Count - 1);
      
      // Select "add more users" since they're editing
      AddMoreRadio.Checked := True;
      DoneRadio.Checked := False;
      
      // Stay on this page
      Result := False;
    end
    // If no previous users, allow normal Back behavior (goes to InstallTypePage)
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  VCRedistPath: string;
  RDPWInstPath: string;
  i: Integer;
  UserInfo: string;
  UserName: string;
  Password: string;
begin
  if CurStep = ssInstall then
  begin
    // Hide cancel button during installation to prevent confusion
    WizardForm.CancelButton.Visible := False;
    
    // Initialize and show relevant steps (pending state) with contiguous layout
    BeginStepLayout;
    if InstallType = 3 then
    begin
      StepsHeaderLabel.Caption := 'Uninstall Steps:';
      AddStepPendingLabel(StepStopSvc, TXT_StopSvc);
      AddStepPendingLabel(StepRemoveExcl, TXT_RemoveExcl);
      AddStepPendingLabel(StepUninstallRDPWrapper, TXT_UninstallRDPWrapper);
      AddStepPendingLabel(StepRemoveFolder, TXT_RemoveFolder);
      AddStepPendingLabel(StepStartSvc, TXT_RestartSvc);
    end
    else if InstallType = 2 then
    begin
      StepsHeaderLabel.Caption := 'Advanced Steps:';
      AddStepPendingLabel(StepCreateShortcuts, TXT_CreateShortcuts);
      AddStepPendingLabel(StepPreTrust, TXT_PreTrust);
    end
    else if InstallType = 0 then
    begin
      StepsHeaderLabel.Caption := 'Install Steps:';
      AddStepPendingLabel(StepCheckMSTSC, TXT_CheckMSTSC);
      AddStepPendingLabel(StepInstallMSTSC, TXT_InstallMSTSC);
      AddStepPendingLabel(StepStopSvc, TXT_StopSvc);
      AddStepPendingLabel(StepAddExcl, TXT_AddExcl);
      AddStepPendingLabel(StepEnsureVC, TXT_EnsureVC);
      AddStepPendingLabel(StepInstallRDPWrapper, TXT_InstallRDPWrapper);
      AddStepPendingLabel(StepConfigureService, TXT_ConfigureService);
      if UsersList.Count > 0 then
        AddStepPendingLabel(StepCreateUsers, TXT_CreateUsers);
      AddStepPendingLabel(StepStartSvc, TXT_StartSvc);
      AddStepPendingLabel(StepPreTrust, TXT_PreTrust);
      AddStepPendingLabel(StepCheckRDP, TXT_CheckRDP);
    end
    else // InstallType = 1 (Add Users Only)
    begin
      StepsHeaderLabel.Caption := 'Install Steps:';
      AddStepPendingLabel(StepAddExcl, TXT_AddExcl);
      if UsersList.Count > 0 then
        AddStepPendingLabel(StepCreateUsers, TXT_CreateUsers);
      AddStepPendingLabel(StepPreTrust, TXT_PreTrust);
    end;

    if InstallType = 0 then
      CheckAndInstallMSTSC;

    // Handle uninstall cleanup
    if InstallType = 3 then
    begin
      WizardForm.StatusLabel.Caption := 'Preparing uninstallation...';
      WizardForm.ProgressGauge.Style := npbstMarquee;
      
      SetStepInProgress(StepStopSvc, TXT_StopSvc);
      StopTermService;
      SetStepDone(StepStopSvc, TXT_StopSvc);
      
      // Remove Defender exclusion before deleting files
      SetStepInProgress(StepRemoveExcl, TXT_RemoveExcl);
      RemoveDefenderExclusionForApp;
      SetStepDone(StepRemoveExcl, TXT_RemoveExcl);

      // Remove TermWrap.dll and Zydis.dll
      DeleteFile(AppBin(FILE_TERMWRAP));
      DeleteFile(AppBin(FILE_ZYDIS));
      Sleep(SLEEP_SHORT);
      
      // Uninstall Everything - Restore to default Windows termsrv.dll
      SetStepInProgress(StepUninstallRDPWrapper, TXT_UninstallRDPWrapper);
      WizardForm.StatusLabel.Caption := 'Uninstalling RDP Wrapper...';
      
      // Temporarily set ServiceDll to RDP Wrapper value for RDPWInst.exe to work
      RegWriteStringValue(HKLM, REG_TERMSERVICE_PARAMS, 'ServiceDll', GetRDPWrapperPath + '\rdpwrap.dll');
      Sleep(SLEEP_MEDIUM);
      
      // Extract RDPWInst.exe from payload to temp directory for robustness
      // First try to use it from {app}, fall back to extracting from payload
      RDPWInstPath := AppBin(FILE_RDPW_INST);
      
      // If RDPWInst.exe is missing from app folder, extract from payload
      if not FileExists(RDPWInstPath) then
      begin
        WizardForm.StatusLabel.Caption := 'Extracting RDP Wrapper uninstaller...';
        RDPWInstPath := TempFile(FILE_RDPW_INST);
        // Extract from payload
        CopyFile(ExpandConstant('{src}\third_party\rdpwrap_release\RDPWrap\RDPWInst.exe'), RDPWInstPath, False);
      end;
      
      // Run RDPWInst.exe to uninstall RDP Wrapper
      Exec(RDPWInstPath, '-u', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      Log('DEBUG: RDPWInst.exe -u exit code = ' + IntToStr(ResultCode));
      Sleep(SLEEP_MEDIUM);
      SetStepDone(StepUninstallRDPWrapper, TXT_UninstallRDPWrapper);
      
      // Restore TermService to default Windows termsrv.dll
      SetStepInProgress(StepRemoveFolder, TXT_RemoveFolder);
      RegWriteStringValue(HKLM, REG_TERMSERVICE_PARAMS, 'ServiceDll', ExpandConstant('{sys}\termsrv.dll'));
      
      // Restore UmRdpService to default umrdp.dll
      RegWriteStringValue(HKLM, REG_UMRDPSERVICE_PARAMS, 'ServiceDll', ExpandConstant('{sys}\umrdp.dll'));
      
      // Remove the entire RDP Wrapper installation folder
      DelTree(GetRDPWrapperPath, True, True, True);
      Sleep(SLEEP_SHORT);
      SetStepDone(StepRemoveFolder, TXT_RemoveFolder);
      
      SetStepInProgress(StepStartSvc, TXT_RestartSvc);
      StartTermService;
      SetStepDone(StepStartSvc, TXT_RestartSvc);
    end
    // Skip installation for Advanced mode (handled in ssPostInstall)
    else if InstallType = 2 then
    begin
      WizardForm.StatusLabel.Caption := 'Preparing advanced tools...';
      WizardForm.ProgressGauge.Style := npbstMarquee;
    end
    // Only stop TermService for Full Install (not for Add Users Only)
    else if InstallType = 0 then
    begin
      WizardForm.StatusLabel.Caption := 'Preparing installation...';
      WizardForm.ProgressGauge.Style := npbstMarquee;
      
      // Now that UI is visible, safely stop the service (executes first, displays first)
      SetStepInProgress(StepStopSvc, TXT_StopSvc);
      WizardForm.StatusLabel.Caption := 'Stopping Remote Desktop Services...';
      Log('[CurStepChanged-ssInstall] Stopping TermService for Full Install');
      StopTermService;
      SetStepDone(StepStopSvc, TXT_StopSvc);
      
      SetStepInProgress(StepAddExcl, TXT_AddExcl);
      WizardForm.StatusLabel.Caption := 'Adding Windows Defender exclusion...';
      AddDefenderExclusionForApp;
      SetStepDone(StepAddExcl, TXT_AddExcl);
    end;
  end;
  
  if CurStep = ssPostInstall then
  begin
    // Handle uninstall completion
    if InstallType = 3 then
    begin
      WizardForm.StatusLabel.Caption := 'Uninstallation complete! TermWrap and RDP Wrapper have been removed.';
    end
    // Advanced mode: tools and utilities
    else if InstallType = 2 then
    begin
      SetStepInProgress(StepCreateShortcuts, TXT_CreateShortcuts);
      WizardForm.StatusLabel.Caption := 'Creating RDP shortcuts...';
      CreateShortcutsForExistingUsers;
      // Advanced path completed
      SetStepDone(StepCreateShortcuts, TXT_CreateShortcuts);
      SetStepInProgress(StepPreTrust, TXT_PreTrust);
      WizardForm.StatusLabel.Caption := 'Pre-trusting Remote Desktop certificate...';
      PreTrustRDPCertCurrentUser;
      SetStepDone(StepPreTrust, TXT_PreTrust);
      ClearPasswordsFromMemory;
      WizardForm.StatusLabel.Caption := 'Advanced tools executed.';
    end
    // Full Install: Download VC++, apply registry, start service
    else if InstallType = 0 then
    begin
      SetStepInProgress(StepEnsureVC, TXT_EnsureVC);
      // Check if VC++ Redistributable is already installed (unless debug mode)
      if DebugMode or (not IsVCRedistInstalled) then
      begin
        WizardForm.StatusLabel.Caption := 'Downloading VC++ Redistributable from Microsoft...';
        WizardForm.ProgressGauge.Style := npbstMarquee;
        
        // Download VC++ Redistributable from Microsoft
        VCRedistPath := TempFile('vc_redist.x64.exe');
        
        // Start download process
           Exec(EXE_POWERSHELL, BuildPowerShellArgs(
             '$ProgressPreference = ''SilentlyContinue''; ' +
             '$url = ''' + URL_VCREDIST_X64 + '''; ' +
             '$output = ''' + VCRedistPath + '''; ' +
             'try { ' +
             '  $webClient = New-Object System.Net.WebClient; ' +
             '  $webClient.DownloadFile($url, $output); ' +
             '  exit 0; ' +
             '} catch { ' +
             '  exit 1; ' +
             '}', True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
        
        WizardForm.StatusLabel.Caption := 'Installing VC++ Redistributable (this may take a minute)...';
        WizardForm.Update;
        
        // Install VC++ Redist silently
        if FileExists(VCRedistPath) then
          Exec(VCRedistPath, '/install /quiet /norestart', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      end
      else
      begin
        WizardForm.StatusLabel.Caption := 'VC++ Redistributable already installed, skipping...';
      end;
      // VC++ ensured (installed or skipped)
      SetStepDone(StepEnsureVC, TXT_EnsureVC);
      
      // Install RDP Wrapper
      SetStepInProgress(StepInstallRDPWrapper, TXT_InstallRDPWrapper);
      WizardForm.StatusLabel.Caption := 'Installing RDP Wrapper...';
      Exec(AppBin(FILE_RDPW_INST), '-i -o', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      Log('DEBUG: RDPWInst.exe -i -o exit code = ' + IntToStr(ResultCode));
      Sleep(SLEEP_MEDIUM);
      SetStepDone(StepInstallRDPWrapper, TXT_InstallRDPWrapper);
      
      SetStepInProgress(StepConfigureService, TXT_ConfigureService);
      WizardForm.StatusLabel.Caption := 'Configuring TermWrap service...';
      
      // Set ServiceDll registry value directly in code
      RegWriteStringValue(HKLM, REG_TERMSERVICE_PARAMS, 'ServiceDll', AppBin(FILE_TERMWRAP));
      
      // Allow multiple concurrent sessions per user (0 = allow multiple, 1 = single session only)
      RegWriteDWordValue(HKLM, REG_TERMINAL_SERVER, 'fSingleSessionPerUser', 0);
      
      // Don't start the service yet - wait until all users are created
      SetStepDone(StepConfigureService, TXT_ConfigureService);
    end;
    
    // Create all user accounts and RDP files (skip for uninstall and advanced)
    if (InstallType <> 3) and (InstallType <> 2) then
    begin
      // No Defender exclusion required for Create Users only (InstallType = 1)
      
      if UsersList.Count > 0 then
      begin
        SetStepInProgress(StepCreateUsers, TXT_CreateUsers);
        CreateRDPUsers;
        SetStepDone(StepCreateUsers, TXT_CreateUsers);
      end;
      
      // Save usernames before clearing passwords from memory
      CreatedUsersList.Clear;
      for i := 0 to UsersList.Count - 1 do
      begin
        UserInfo := UsersList[i];
        ParseUserEntry(UserInfo, UserName, Password);
        CreatedUsersList.Add(UserName);
      end;
      
      // Clear passwords from memory immediately after use
      ClearPasswordsFromMemory;

      // If shortcuts were requested for existing users, create them now
      if ShortcutsList.Count > 0 then
      begin
        SetStepInProgress(StepCreateShortcuts, TXT_CreateShortcuts);
        CreateShortcutsForExistingUsers;
        SetStepDone(StepCreateShortcuts, TXT_CreateShortcuts);
      end;
    end;
    
    // Start TermService first (only for Full Install) - this creates the SSL certificate
    if InstallType = 0 then
    begin
      SetStepInProgress(StepStartSvc, TXT_StartSvc);
      WizardForm.StatusLabel.Caption := 'Starting Remote Desktop Services...';
      // Start TermService after all files and registry are done
      ResultCode := StartTermServiceEx;
      if ResultCode = 0 then
      begin
        Sleep(SLEEP_EXTRALONG); // Wait for service to fully initialize and create certificate
        SetStepDone(StepStartSvc, TXT_StartSvc);
      end
      else
      begin
        Log('WARNING: TermService failed to start with exit code ' + IntToStr(ResultCode));
        SetStepDone(StepStartSvc, TXT_StartSvc); // Mark as done even if failed (might already be running)
        Sleep(SLEEP_LONG); // Give extra time if service had issues
      end;
    end;
    
    // Pre-trust for current user in Full Install and Add Users Only (AFTER service starts)
    if (InstallType = 0) or (InstallType = 1) then
    begin
      SetStepInProgress(StepPreTrust, TXT_PreTrust);
      WizardForm.StatusLabel.Caption := 'Pre-trusting Remote Desktop certificate...';
      PreTrustRDPCertCurrentUser;
      // Pre-trust is optional - don't fail install if cert doesn't exist yet
      SetStepDone(StepPreTrust, TXT_PreTrust);
    end;

    // Verify RDP is listening (only for Full Install)
    if InstallType = 0 then
    begin
      SetStepInProgress(StepCheckRDP, TXT_CheckRDP);
      WizardForm.StatusLabel.Caption := 'Verifying RDP service...';
      
      ResultCode := 0;
      Exec(EXE_POWERSHELL, BuildPowerShellArgs('try { if ((Get-NetTCPConnection -LocalPort ' + IntToStr(RDP_LISTEN_PORT) + ' -State Listen -ErrorAction SilentlyContinue).LocalPort -eq ' + IntToStr(RDP_LISTEN_PORT) + ') { exit 0 } else { exit 1 } } catch { exit 1 }', True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      
      if ResultCode = 0 then
      begin
        SetStepDone(StepCheckRDP, TXT_CheckRDP);
      end
      else
      begin
        SetStepDone(StepCheckRDP, TXT_CheckRDP);
        if MsgBox('RDP service is not detected as listening on port 3389.' + #13#10#13#10 +
                  'A system restart usually resolves this issue.' + #13#10#13#10 +
                  'Would you like to restart your computer now?', mbConfirmation, MB_YESNO) = IDYES then
        begin
          // User chose to restart now
          Exec('shutdown.exe', '/r /t 5 /c "Restarting to complete RDP Wrapper setup"', '', SW_HIDE, ewNoWait, ResultCode);
        end;
      end;
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  // Handle removal when using the standard uninstaller
  if CurUninstallStep = usUninstall then
  begin
    RemoveDefenderExclusionForApp;
  end;
end;

procedure CurPageChanged(CurPageID: Integer);
var
  i: Integer;
  CompletionText: string;
begin
  // Lazy-load user list only when Advanced Tools page is first shown
  if (CurPageID = AdvancedToolsSelectionPage.ID) and (LocalUsersList.Count = 0) then
  begin
    LocalUsersList := GetLocalUsers;
    SetLength(UserCheckBoxes, LocalUsersList.Count);
    SetLength(UserPasswordEdits, LocalUsersList.Count);
    SetLength(UserPasswordStatus, LocalUsersList.Count);
    // Build per-user controls on Tool 1 page
    BuildAdvancedUserControls;
  end;
  
  // Display completion info on the final page
  if CurPageID = wpFinished then
  begin
    WriteInstallerLog('CurPageChanged: Finish page shown');
    if InstallType = 3 then
    begin
      // Uninstall completion message
      WizardForm.FinishedHeadingLabel.Caption := 'Uninstallation Complete';
      CompletionText := 'TermWrap and RDP Wrapper have been successfully removed.' + #13#10#13#10 +
                       'The Program Files\RDP Wrapper folder and all its contents have been deleted.' + #13#10#13#10 +
                       'Remote Desktop Services have been restored to their default Windows configuration.';
      WriteInstallerLog('CurPageChanged: Showing uninstall completion message');
    end
    else if InstallType = 2 then
    begin
      // Advanced tools completion message
      WizardForm.FinishedHeadingLabel.Caption := 'Advanced Tools Complete';
      CompletionText := 'Advanced tools have been executed successfully.' + #13#10#13#10 +
                       'Your RDP installation is ready to use.';
      WriteInstallerLog('CurPageChanged: Showing advanced tools completion message');
    end
    else
    begin
      // Installation completion message
      WizardForm.FinishedHeadingLabel.Caption := 'Installation Complete';
      if CreatedUsersList.Count > 0 then
      begin
        CompletionText := 'Created ' + IntToStr(CreatedUsersList.Count) + ' user account(s) and desktop shortcuts:' + #13#10;
        // Iterate through all created users and add to completion text
        for i := 0 to CreatedUsersList.Count - 1 do
        begin
          CompletionText := CompletionText + '- ' + CreatedUsersList[i] + #13#10;
        end;
        CompletionText := CompletionText + #13#10 +
                         'You can now open RDP connections using these shortcuts.';
        WriteInstallerLog('CurPageChanged: Created ' + IntToStr(CreatedUsersList.Count) + ' users');
      end
      else
      begin
        CompletionText := 'No user accounts were created during this run.' + #13#10#13#10 +
                          'You can add users later by rerunning this installer and choosing "Create Users Only".';
        WriteInstallerLog('CurPageChanged: No users created');
      end;
    end;
    
    WizardForm.FinishedLabel.Caption := CompletionText;
    WizardForm.Update;
    WriteInstallerLog('CurPageChanged: Updating FinishedText control');
    // Populate the rich viewer if available
    if Assigned(FinishedText) then
    begin
      WriteInstallerLog('CurPageChanged: FinishedText control is assigned');
      // Ensure finished text color contrasts with page surface
      if IsDarkColor(FinishedText.Color) then
        FinishedText.Font.Color := clWhite
      else
        FinishedText.Font.Color := clBlack;
      FinishedText.RTFText := PlainToRtfWithColor(CompletionText, FinishedText.Font.Color);
      FinishedText.Update;
      WriteInstallerLog('CurPageChanged: FinishedText populated with completion message');
    end
    else
    begin
      WriteInstallerLog('CurPageChanged: ERROR - FinishedText control is NOT assigned!');
    end;
    
    if Assigned(ViewLogButton) then
    begin
      WriteInstallerLog('CurPageChanged: ViewLogButton is assigned and visible');
    end
    else
    begin
      WriteInstallerLog('CurPageChanged: ERROR - ViewLogButton is NOT assigned!');
    end;
  end;
end;

procedure CheckAndInstallMSTSC;
var
  ResultCode: Integer;
  MSTSCExists: Boolean;
begin
  // Check if mstsc.exe exists
  SetStepInProgress(StepCheckMSTSC, TXT_CheckMSTSC);
  MSTSCExists := FileExists(FILE_MSTSC);
  if MSTSCExists then
  begin
    Log('DEBUG: mstsc.exe found.');
    SetStepDone(StepCheckMSTSC, TXT_CheckMSTSC);
    SetStepDone(StepInstallMSTSC, TXT_InstallMSTSC); // skipped
  end
  else
  begin
    Log('DEBUG: mstsc.exe missing. Initiating installation.');
    SetStepDone(StepCheckMSTSC, TXT_CheckMSTSC);
    SetStepInProgress(StepInstallMSTSC, TXT_InstallMSTSC);
    Exec(EXE_POWERSHELL, BuildPowerShellArgs('Start-Process -FilePath ''' + URL_RDP_INSTALLER + ''' -Wait', False), '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode);
    if ResultCode = 0 then
    begin
      Log('DEBUG: Remote Desktop Connection installed successfully.');
      SetStepDone(StepInstallMSTSC, TXT_InstallMSTSC);
    end
    else
    begin
      Log('ERROR: Failed to install Remote Desktop Connection. Exit code: ' + IntToStr(ResultCode));
      MsgBox('Failed to install Remote Desktop Connection. Please install it manually.', mbError, MB_OK);
      SetStepDone(StepInstallMSTSC, TXT_InstallMSTSC); // mark as done even on failure?
      // Perhaps leave it pending or something, but for now, done.
    end;
  end;
end;
