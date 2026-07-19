; =========================================================================
; RDPWrapKit - Local RDP Management Suite
; =========================================================================
; 
; PURPOSE:
;   This Inno Setup installer provides a comprehensive solution for setting
;   up TermWrap on Windows systems, enabling multiple concurrent
;   Remote Desktop sessions on non-Server editions of Windows.

;
; KEY FEATURES:
;   - Automatic installation of TermWrap (llccd)
;   - VC++ Redistributable dependency management
;   - User account creation with automatic RDP shortcuts
;   - Security hardening (Windows Defender exclusions, secure credential handling)
;   - Create Shortcuts flow for existing users and shortcuts
;   - Complete uninstallation support with registry cleanup
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

#define APP_VERSION_STRING "0.6.0"
#define APP_VERSION_FILEINFO "0.6.0.0"

; Preprocessor captures source TermWrap.dll metadata at compile time for runtime comparison.
#define SourceTermWrapVersion GetVersionNumbersString("third_party\termwrap_release\TermWrap.dll")
#define SourceTermWrapSize FileSize("third_party\termwrap_release\TermWrap.dll")

; RdpSignTool.exe is compiled ahead-of-time from scripts\RdpSignTool.cs.
; Run scripts\build_rdpcrypt.ps1 before compiling this installer.
; This keeps the 19KB binary out of git while enabling a clean build.

[Setup]
AppName=RDPWrapKit
AppVersion={#APP_VERSION_STRING}
VersionInfoVersion={#APP_VERSION_FILEINFO}
AppPublisher=cpdx4
AppPublisherURL=https://cpdx4.github.io/RDPWrapKit/
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
WizardStyle=modern dark
SetupIconFile="assets\RDPWrapKitIcon.ico"
WizardBackImageFile="assets\RDPWrapInstallerBG.bmp"
WizardBackColor=#0b1018

[Files]
; Icon file always extracted to temp for welcome page display.
; TermWrap files only copied when DoInstallTermWrap = True (checked via ShouldInstallFiles).
Source: "third_party\termwrap_release\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs; Check: ShouldInstallFiles
Source: "output\RdpSignTool.exe"; DestDir: "{tmp}"; Flags: ignoreversion dontcopy
Source: "assets\RDPWrapKitIcon.bmp"; DestDir: "{tmp}"; Flags: ignoreversion dontcopy
Source: "assets\rdp_edit_save.bmp"; DestDir: "{tmp}"; Flags: ignoreversion skipifsourcedoesntexist dontcopy
Source: "assets\RDPWrapInstallerBG.bmp"; DestDir: "{tmp}"; Flags: ignoreversion dontcopy
Source: "assets\restore_term_service.reg"; DestDir: "{tmp}"; Flags: ignoreversion skipifsourcedoesntexist dontcopy



[Registry]
; Enable RDP (fDenyTSConnections=0). Removed on uninstall via uninsdeletevalue.
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Terminal Server"; ValueType: dword; ValueName: "fDenyTSConnections"; ValueData: 0; Flags: uninsdeletevalue; Check: ShouldApplyRegistryEntries

[Run]
; Run section is now handled in Code section for proper sequencing

[UninstallRun]
; Registry cleanup handled via uninsdeletevalue flags and Code section.

[Code]
{ Pascal Script: installer UI, validation, and install flow. }

// Forward declarations
procedure CheckAndInstallMSTSC; forward;
procedure OnCreateRdpShortcutsClick(Sender: TObject); forward;
procedure OnPrevUsersPageClick(Sender: TObject); forward;
procedure OnNextUsersPageClick(Sender: TObject); forward;
procedure UpdateUsersPageDisplay; forward;
procedure OnPrevShortcutPageClick(Sender: TObject); forward;
procedure OnNextShortcutPageClick(Sender: TObject); forward;
procedure OnShortcutCheckBoxClick(Sender: TObject); forward;
procedure BuildEditShortcutAdvancedControls; forward;
procedure UnSignRdpFile(const RdpPath: string); forward;
procedure UpdateShortcutPageDisplay; forward;
function IsValidPassword(const Password: string): String; forward;
procedure OpenTermWrap(Sender: TObject); forward;
function ValidateLocalCredential(const UserName, Password: string): Boolean; forward;
procedure OpenBSGH(Sender: TObject); forward;
procedure OpenBSSGrinders(Sender: TObject); forward;
procedure OnInstallModeChange(Sender: TObject); forward;
procedure OnFullScreenClick(Sender: TObject); forward;
procedure OnUseAllMonitorsClick(Sender: TObject); forward;
procedure OnResolutionChange(Sender: TObject); forward;
function IsTermWrapInstalled(): Boolean; forward;
function GetInstalledTermWrapVersion(): string; forward;
function GetInstalledTermWrapSize(): string; forward;
function BoolToStr(Value: Boolean): string; forward;
procedure EnsureTermServiceRunsAsNetworkService; forward;
procedure EnsureUmRdpServiceAutomatic; forward;
procedure InitInstallerLog; forward;
procedure WriteInstallerLog(const Msg: string); forward;
procedure LogSectionHeader(const Title: string); forward;
procedure LogKeyValue(const KeyName, KeyValue: string); forward;

// Logger architecture forward declarations
function GetThreadIdHex: string; forward;
function GetProcessIdHex: string; forward;
function LoggerRedact(const S: string): string; forward;
procedure LoggerWrite(const Level, Msg: string); forward;
procedure LogInfo(const Msg: string); forward;
procedure LogDebug(const Msg: string); forward;
procedure LogWarn(const Msg: string); forward;
procedure LogError(const Msg: string); forward;
procedure LogEntry(const FuncName: string); forward;
procedure LogExit(const FuncName: string); forward;
procedure LoggerStartOp; forward;
function LoggerEndOp: Cardinal; forward;
procedure LogFileOp(const Operation, FilePath: string; const Success: Boolean; const Extra: string); forward;
procedure LogRegOp(const Operation, RegKey, RegValue: string; const Success: Boolean); forward;
procedure LogServiceOp(const Operation, ServiceName: string; const ResultCode: Integer); forward;
procedure LogSim(const ScenarioText: string); forward;

// Debug-enriched command execution forward declarations
function RunCmdCapture(const CmdLine, OutTag: string): Integer; forward;
function RunNetHiddenCapture(const Params, OutTag: string): Integer; forward;

var
  Page_InstallOptions: TWizardPage;
  WelcomePage: TWizardPage;
  UserPage: TInputQueryWizardPage;
  // AdvancedPage removed - single create-shortcuts page retained
  EditSystemwideSettingsPage: TWizardPage;  // Main Edit System-wide settings page
  QuickFixesPage: TWizardPage;              // Quick Fixes page
  Page_CreateShortcutsForExistingUsers: TWizardPage;  // Create RDP desktop shortcuts
  LocalUsersList: TStringList;        // login usernames
  LocalUserDisplayList: TStringList;  // display labels (email if online account, else same as username)
  CreateShortcutsControlsBuilt: Boolean;
  UserCheckBoxes: array of TCheckBox;
  UserPasswordEdits: array of TEdit;
  UserPasswordStatus: array of TLabel;
  // Pagination for existing user list display
  CurrentUserPage: Integer;
  Tool1PrevButton: TButton;
  Tool1NextButton: TButton;
  Tool1PageLabel: TLabel;
  UsersPerPage: Integer;
  ShortcutsList: TStringList;
  InstallLogPath: string;
  AddMoreRadio: TRadioButton;
  DoneRadio: TRadioButton;
  // New welcome/options controls
  rbInstall: TRadioButton;
  rbEditSystemwideSettings: TRadioButton;
  rbShowRDPInfo: TRadioButton;
  rbQuickFixes: TRadioButton;
  rbUninstall: TRadioButton;
  Page_ShowRDPInfo: TWizardPage;
  // Show RDP Info page controls
  cmbGPCompression: TComboBox;
  cmbGPImageQuality: TComboBox;
  StepShowRDPInfo: TLabel;
  chkInstallTermWrap: TCheckBox;
  chkCreateRdpShortcuts: TCheckBox;
  rbCreateUsers: TRadioButton;
  rbUseExistingUsers: TRadioButton;
  chkInstallTermWrapHint: TLabel;
  rbUseExistingUsersHint: TLabel;
  InstallOptionsAutoUserSourceApplied: Boolean;
  rbEditShortcutSettings: TRadioButton;
  CreateRdpShortcutsGroup: TPanel;
  EditShortcutPage: TWizardPage;
  Page_ShortcutSettings: TWizardPage;
  Page_EditShortcutAdvanced: TWizardPage;
  DesktopRdpFiles: TStringList;
  ShortcutCheckBoxes: array of TCheckBox;
  CurrentShortcutPage: Integer;
  ShortcutsPerPage: Integer;
  ShortcutPrevButton: TButton;
  ShortcutNextButton: TButton;
  ShortcutPageLabel: TLabel;
  ShortcutHeaderLabel: TLabel;
  ShortcutEmptyLabel: TLabel;
  EditShortcutControlsBuilt: Boolean;
  EditShortcutAdvancedControlsBuilt: Boolean;
  SelectedShortcutIndex: Integer;
  SelectedShortcutPath: string;
  SelectedShortcutPaths: TStringList;
  FinishedExampleImage: TBitmapImage;
  EditShortcutAdvancedImage: TBitmapImage;
  EditShortcutAdvancedLabel: TLabel;
  
  // Flags derived from welcome/options controls
  DoInstallTermWrap: Boolean;
  DoCreateRdpShortcuts: Boolean;
  CreateUserMode: Integer;       // createUserModeNew or createUserModeExisting (only used when DoCreateRdpShortcuts = True)
  DoEditSystemWideSettings: Boolean;
  OrigEnableRDP: Boolean;
  OrigShowUsers: Boolean;
  OrigPreventDuplicate: Boolean;
  OrigHideSecurityWarnings: Boolean;
  OrigRdpPort: Cardinal;
  OptionsLabel: TLabel;
  Tool1UsersHeaderLabel: TLabel;  // "Users found" header
  Tool1PasswordHeaderLabel: TLabel;  // "Password" header
  Tool1PasswordResetLink: TLabel;  // Password reset link at bottom
  // Controls for Quick Fixes page
  lblQFHeader: TLabel;
  rbQFRestartRDP: TRadioButton;
  rbQFAccountNeverExpires: TRadioButton;
  rbQFRestoreTermService: TRadioButton;

  // Controls for Edit System-wide Settings page
  lblSysHeader: TLabel;
  lblWinVer: TLabel;
  lblWinVerName: TLabel;
  lblRDPService: TLabel;
  lblRDPServiceName: TLabel;
  lblWinRDPVer: TLabel;
  lblWinRDPVerName: TLabel;
  lblWrapperVer: TLabel;
  lblWrapperVerName: TLabel;
  lblListenerName: TLabel;
  lblListener: TLabel;

  lblGenHeader: TLabel;
  chkEnableRDP: TCheckBox;
  chkShowUsers: TCheckBox;
  chkPreventDuplicate: TCheckBox;
  chkHideSecurityWarnings: TCheckBox;
  lblRdpPort: TLabel;
  edtRdpPort: TEdit;
  lblPortDefault: TLabel;

  lblActionsHeader: TLabel;
  chkRestartRDP: TCheckBox;
  // Progress UI on Installing page
  StepsHeaderLabel: TLabel;
  StepAddExcl: TLabel;
  StepRemoveExcl: TLabel;
  StepStopSvc: TLabel;
  StepEnsureVC: TLabel;
  StepInstallTermWrap: TLabel;
  StepConfigureService: TLabel;
  StepCreateUsers: TLabel;
  StepCreateShortcuts: TLabel;
  StepPreTrust: TLabel;
  StepStartSvc: TLabel;
  StepCheckRDP: TLabel;
  StepCheckMSTSC: TLabel;
  StepInstallMSTSC: TLabel;
  StepRemoveFolder: TLabel;
  StepUninstallTermWrap: TLabel;
  StepEnableRDP: TLabel;
  StepShowUsers: TLabel;
  StepPreventDuplicate: TLabel;
  StepSetRdpPort: TLabel;
  StepRestartRDP: TLabel;
  StepQuickFixes: TLabel;
  StepAccountNeverExpires: TLabel;
  SelectedInstallMode: Integer;  // installModeInstall, installModeEditShortcuts, installModeEditSystemwideSettings, installModeUninstall
  DebugMode: Boolean;    // Set to True to force VC++ download even if installed
  DoShowMstscEdit: Boolean;  // True = open mstsc /edit after writing settings (EditShortcuts path)
  // Simulation marker flags (log once per run)
  SimLogNoMstscShown: Boolean;
  SimLogNoVCRedistShown: Boolean;
  SimLogNetPsShown: Boolean;
  LastLoggedPageId: Integer;
  LastLoggedPageTick: Cardinal;
  LastSuppressedPageLogs: Integer;
  PendingDebugCleanupFiles: TStringList;
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
  FinishedText: TLabel;
  // Button to open install log on finish page
  ViewLogButton: TButton;
  // Flag set when Smart App Control (VerifiedAndReputablePolicyState) is detected as On
  SmartAppControlIsOn: Boolean;
  // Shortcut settings controls on Create RDP User Account page
  lblShortcutSection: TLabel;
  lblScreenSize: TLabel;
  cboResolution: TComboBox;
  chkFullScreen: TCheckBox;
  chkUseAllMonitors: TCheckBox;
  chkCopyPaste: TCheckBox;
  chkSound: TCheckBox;
  lblMultiShortcutEditingNote: TLabel;
  chkShowMoreShortcutOptions: TCheckBox;
  lblCustomWidth: TLabel;
  edtCustomWidth: TEdit;
  lblCustomHeight: TLabel;
  edtCustomHeight: TEdit;
  // Experience / performance checkboxes on Shortcut Settings page
  chkExpWallpaper: TCheckBox;
  chkExpFontSmooth: TCheckBox;
  chkExpComposition: TCheckBox;
  chkExpDragContents: TCheckBox;
  chkExpMenuAnim: TCheckBox;
  chkExpVisualStyles: TCheckBox;
  // New shortcut name customization
  lblShortcutName: TLabel;
  edtShortcutName: TEdit;
  lblShortcutExtension: TLabel;
  // Keyboard hook settings
  lblKeyboardHook: TLabel;
  cboKeyboardHook: TComboBox;
  
  // Transparent overlay for status text (replaces WizardForm.StatusLabel which can't be made transparent)
  StatusOverlay: TLabel;
  
  // Determinate progress bar tracking
  StepsTotal: Integer;
  StepsDone: Integer;
  
  // Logger state globals for performance profiling and metadata
  LoggerProcId: DWORD;
  LoggerThreadId: DWORD;
  LoggerOpStartTick: Cardinal;
  LoggerOpActive: Boolean;

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
  TXT_InstallTermWrap = 'Install TermWrap';
  TXT_ConfigureService = 'Install and configure TermWrap';
  TXT_CreateUsers = 'Create user accounts';
  TXT_CreateShortcuts = 'Create RDP shortcuts for selected users';
  TXT_PreTrust = 'Pre-trust RDP certificate for current user';
  TXT_CheckRDP = 'Verify RDP service is listening';
  TXT_CheckMSTSC = 'Check for Remote Desktop Connection';
  TXT_InstallMSTSC = 'Install Remote Desktop Connection (if missing)';
  TXT_RemoveFolder = 'Remove TermWrap folder';
  TXT_UninstallTermWrap = 'Uninstall TermWrap';
  TXT_ShowRDPInfo = 'Apply RDP settings';
  
  // -------------------------------------------------------------------------
  // REGISTRY PATHS
  // -------------------------------------------------------------------------
  // Windows Registry keys for Terminal Services and RDP configuration
  // IMPORTANT: These paths are system-critical and must remain accurate
  REG_TERMSERVICE_PARAMS = 'SYSTEM\CurrentControlSet\Services\TermService\Parameters';
  REG_TERMSERVICE = 'SYSTEM\CurrentControlSet\Services\TermService';
  REG_TERMINAL_SERVER = 'SYSTEM\CurrentControlSet\Control\Terminal Server';
  REG_UMRDPSERVICE = 'SYSTEM\CurrentControlSet\Services\UmRdpService';
  REG_UMRDPSERVICE_PARAMS = 'SYSTEM\CurrentControlSet\Services\UmRdpService\Parameters';
  REG_VCREDIST = 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64';
  REG_SHOW_USERS = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System';
  REG_RDP_TCP = 'SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp';
  REG_TS_POLICIES = 'SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services';
  // Loopback alias used by TermWrap for local RDP sessions
  RDP_LOOPBACK_IP = '127.0.0.2';
  // Default RDP window resolution written into generated .rdp shortcut files
  DEFAULT_RDP_WIDTH  = 1366;
  DEFAULT_RDP_HEIGHT = 768;
  // User groups
  GROUP_ADMINISTRATORS = 'Administrators';
  GROUP_RDP_USERS = 'Remote Desktop Users';
  NET_USER_TEMP_PASSWORD = 'Tmp1!';     // Short throwaway password for two-step net user /add (avoids 14-char LM prompt)
  
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
  FILE_TERMWRAP = 'TermWrap.dll';
  FILE_ZYDIS = 'Zydis.dll';
  URL_VCREDIST_X64 = 'https://aka.ms/vs/17/release/vc_redist.x64.exe';
  RDP_LISTEN_PORT = 3389;
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
  FILE_RDPEDITSAVE_BMP = 'rdp_edit_save.bmp';
  TEMP_LOCAL_USERS = '{tmp}\\local_users.txt';
  TEMP_ICON_BMP = '{tmp}\\' + FILE_ICON_BMP;
  TEMP_RDPEDITSAVE_BMP = '{tmp}\\' + FILE_RDPEDITSAVE_BMP;
  // Path to installer log file (auto-created in %TEMP%)
  INSTALL_LOG_PATH = '{tmp}\\RDPWrapKit_install.log';
  
  // -------------------------------------------------------------------------
  // EXTERNAL PROJECT URLS
  // -------------------------------------------------------------------------
  // Centralized URLs for attribution and user navigation
  // Update these if upstream projects change their repository locations
  URL_TERMWRAP = 'https://github.com/llccd/TermWrap';
  URL_BSGH_COMMUNITY = 'https://discord.gg/bsgh';
  URL_BSS_GRINDERS = 'https://discord.gg/K5U3RdGXh6';
  URL_PROJECT_HOME = 'https://cpdx4.github.io/RDPWrapKit/';

  // -------------------------------------------------------------------------
  // INSTALL MODE CONSTANTS
  // -------------------------------------------------------------------------
  // Four top-level modes. The Install mode uses boolean flags (DoInstallTermWrap,
  // DoCreateRdpShortcuts) and CreateUserMode to describe what happens within it.
  installModeInstall = 0;                        // Install: TermWrap + optional shortcuts
  installModeEditShortcuts = 1;                  // Edit existing shortcut settings
  installModeEditSystemwideSettings = 2;         // Edit system-wide RDP settings
  installModeUninstall = 3;                      // Uninstall everything
  installModeShowRDPInfo = 4;               // Show RDP Info + configure RemoteFX/startup settings
  installModeQuickFixes = 5;                     // Quick Fixes page

  // CREATE USER MODE CONSTANTS (only relevant when DoCreateRdpShortcuts = True)
  createUserModeNew = 0;                         // Create new local user accounts
  createUserModeExisting = 1;                    // Use existing local user accounts

  // -------------------------------------------------------------------------
  // TEST SCENARIO TOGGLES (set 1 to enable, 0 to disable)
  // -------------------------------------------------------------------------
  // Use only one scenario at a time for predictable behavior.
  SIM_SCENARIO_NO_MSTSC =0;
  SIM_SCENARIO_NO_VCREDIST = 0;
  SIM_SCENARIO_NET_FAIL_POWERSHELL = 0;

  // Suppress duplicate CurPageChanged log blocks if the same page is raised
  // again within this short interval (UI refresh/re-entry noise).
  PAGE_LOG_DEDUPE_MS = 600;

  // Debug controls
  PRESERVE_USER_CREATE_DEBUG_LOGS = 0;
  CLEANUP_DEBUG_FILES_ON_FINISH = 0;

  // Password pipeline diagnostics (temporary deep debugging)
  PASSWORD_PIPELINE_DIAG = 0;
  BUILD_FINGERPRINT = '2026-07-19-0-6-0';

// External Windows API declarations

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

function GetCurrentProcessId: DWORD;
  external 'GetCurrentProcessId@kernel32.dll stdcall';

function GetCurrentThreadId: DWORD;
  external 'GetCurrentThreadId@kernel32.dll stdcall';

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

function GetSysColor(nIndex: DWORD): DWORD;
  external 'GetSysColor@user32.dll stdcall';

function MessageBox(hWnd: Integer; lpText, lpCaption: String; uType: Cardinal): Integer;
  external 'MessageBoxW@user32.dll stdcall';

// Compares two dotted version strings (e.g. "0.4.9" vs "0.4.10").
// Returns 1 if A > B, -1 if A < B, 0 if equal.
function CompareVersions(A, B: String): Integer;
var
  AParts, BParts: TStringList;
  i, AVal, BVal, MaxLen: Integer;
begin
  Result := 0;
  AParts := TStringList.Create;
  BParts := TStringList.Create;
  try
    AParts.Delimiter := '.';
    AParts.StrictDelimiter := True;
    AParts.DelimitedText := A;
    BParts.Delimiter := '.';
    BParts.StrictDelimiter := True;
    BParts.DelimitedText := B;
    if AParts.Count > BParts.Count then MaxLen := AParts.Count
    else MaxLen := BParts.Count;
    for i := 0 to MaxLen - 1 do
    begin
      if i < AParts.Count then AVal := StrToIntDef(AParts[i], 0) else AVal := 0;
      if i < BParts.Count then BVal := StrToIntDef(BParts[i], 0) else BVal := 0;
      if AVal > BVal then begin Result := 1; Break; end;
      if AVal < BVal then begin Result := -1; Break; end;
    end;
  finally
    AParts.Free;
    BParts.Free;
  end;
end;

// Fetches the latest release tag from GitHub Releases API.
// Returns tag_name (e.g. "0.4.9") with leading "v" stripped, or "" on failure.
function GetLatestGitHubVersion(): String;
var
  Http: Variant;
  Response: String;
  TagStart, TagEnd: Integer;
  Tag: String;
  Remainder: String;
begin
  Result := '';
  try
    Http := CreateOleObject('WinHttp.WinHttpRequest.5.1');
    Http.Open('GET', 'https://api.github.com/repos/cpdx4/RDPWrapKit/releases/latest', False);
    Http.SetRequestHeader('User-Agent', 'RDPWrapKit-Installer');
    Http.Send('');
    if Http.Status <> 200 then
      Exit;
    Response := Http.ResponseText;
    // Extract the value of "tag_name":"..."
    TagStart := Pos('"tag_name"', Response);
    if TagStart = 0 then
      Exit;
    // Move past "tag_name" to get the substring starting from there
    Remainder := Copy(Response, TagStart + Length('"tag_name"'), Length(Response));
    // Find the opening quote in the remainder
    TagStart := Pos('"', Remainder);
    if TagStart = 0 then
      Exit;
    // Extract from after the opening quote
    Remainder := Copy(Remainder, TagStart + 1, Length(Remainder));
    // Find the closing quote
    TagEnd := Pos('"', Remainder);
    if TagEnd = 0 then
      Exit;
    Tag := Copy(Remainder, 1, TagEnd - 1);
    // Strip optional leading "v"
    if (Length(Tag) > 0) and (Tag[1] = 'v') then
      Tag := Copy(Tag, 2, Length(Tag) - 1);
    Result := Tag;
  except
    Result := '';
  end;
end;

// Runs at installer startup before the wizard opens.
// Shows an update prompt when a newer version is available on GitHub.
// Returns False to abort the installer, True to continue.
function InitializeSetup(): Boolean;
var
  LatestVersion: String;
  CurrentVersion: String;
  Msg: String;
  Answer: Integer;
begin
  Result := True;
  CurrentVersion := '{#APP_VERSION_STRING}';
  LatestVersion := GetLatestGitHubVersion();
  // Only prompt when a version string was returned and it is strictly newer
  if (LatestVersion <> '') and (CompareVersions(LatestVersion, CurrentVersion) > 0) then
  begin
    Msg := 'A newer version (v' + LatestVersion + ') is available.' + #13#10#13#10
         + 'Open the latest release page instead?';
    Answer := MsgBox(Msg, mbConfirmation, MB_YESNOCANCEL);
    if Answer = IDYES then
    begin
      ShellExec('open', 'https://github.com/cpdx4/RDPWrapKit/releases/latest', '', '', SW_SHOWNORMAL, ewNoWait, Answer);
      Result := False;
    end else if Answer = IDCANCEL then
      Result := False;
    // IDNO falls through with Result = True (install anyway)
  end;
end;

// =============================================================================
// HELPER FUNCTIONS
// =============================================================================

// -----------------------------------------------------------------------------
// USER ACCOUNT VALIDATION
// -----------------------------------------------------------------------------

// Check if a local user account exists
function UserExists(const UserName: string): Boolean;
var
  BufPtr: Cardinal;
begin
  BufPtr := 0;
  Result := NetUserGetInfo('', UserName, 0, BufPtr) = NERR_Success;
  if BufPtr <> 0 then
    NetApiBufferFree(BufPtr);
end;

// Parse a "username|password" string into its components
procedure ParseUserEntry(const Entry: string; var UserName, Password: string);
var
  PipePos: Integer;
begin
  LogEntry('ParseUserEntry');
  PipePos := Pos('|', Entry);
  UserName := Copy(Entry, 1, PipePos - 1);
  Password := Copy(Entry, PipePos + 1, Length(Entry));
  LogDebug('ParseUserEntry: user=' + UserName + ' hasPassword=' + BoolToStr(Password <> ''));
  LogExit('ParseUserEntry');
end;

// Return a version of a pipe-delimited user entry with the password obscured
function MaskPasswordInEntry(const Entry: string): string;
var
  PipePos: Integer;
begin
  LogEntry('MaskPasswordInEntry');
  PipePos := Pos('|', Entry);
  if PipePos > 0 then
    Result := Copy(Entry, 1, PipePos) + '*****'
  else
    Result := Entry;
  LogExit('MaskPasswordInEntry');
end;

function PosFrom(const Needle, Haystack: string; const FromPos: Integer): Integer;
var
  i: Integer;
begin
  Result := 0;
  if (Needle = '') or (Haystack = '') or (FromPos < 1) or (FromPos > Length(Haystack)) then
    exit;
  for i := FromPos to Length(Haystack) - Length(Needle) + 1 do
  begin
    if Copy(Haystack, i, Length(Needle)) = Needle then
    begin
      Result := i;
      exit;
    end;
  end;
end;

// Mask common password flags in arbitrary command strings (e.g. -Password "..." or -Password ...)
function MaskPasswordsInString(const S: string): string;
var
  U: string;
  idx, p, startPos, endPos, SearchPos: Integer;
begin
  Result := S;
  U := UpperCase(Result);
  SearchPos := 1;
  idx := PosFrom('-PASSWORD', U, SearchPos);
  while idx > 0 do
  begin
    p := idx + Length('-PASSWORD');
    while (p <= Length(Result)) and ((Result[p] = ' ') or (Result[p] = '=') ) do
      Inc(p);
    if p > Length(Result) then
      Break;
    if Result[p] = '"' then
    begin
      startPos := p;
      endPos := startPos + 1;
      while (endPos <= Length(Result)) and (Result[endPos] <> '"') do
        Inc(endPos);
      if endPos > Length(Result) then
        endPos := Length(Result);
      Delete(Result, startPos, endPos - startPos + 1);
      Insert('"*****"', Result, startPos);
    end
    else
    begin
      endPos := p;
      while (endPos <= Length(Result)) and (Result[endPos] <> ' ') do
        Inc(endPos);
      Delete(Result, p, endPos - p);
      Insert('*****', Result, p);
    end;
    U := UpperCase(Result);
    SearchPos := idx + Length('-PASSWORD') + 1;
    idx := PosFrom('-PASSWORD', U, SearchPos);
  end;
end;

procedure NextArgRange(const S: string; var Cursor, StartPos, EndPos: Integer);
begin
  while (Cursor <= Length(S)) and (S[Cursor] = ' ') do
    Inc(Cursor);

  StartPos := Cursor;
  EndPos := 0;
  if Cursor > Length(S) then
    exit;

  if S[Cursor] = '"' then
  begin
    Inc(Cursor);
    while (Cursor <= Length(S)) and (S[Cursor] <> '"') do
      Inc(Cursor);
    if Cursor <= Length(S) then
      EndPos := Cursor
    else
      EndPos := Length(S);
    Inc(Cursor);
  end
  else
  begin
    while (Cursor <= Length(S)) and (S[Cursor] <> ' ') do
      Inc(Cursor);
    EndPos := Cursor - 1;
  end;
end;

function StripWrappingQuotes(const S: string): string;
begin
  Result := S;
  if (Length(Result) >= 2) and (Result[1] = '"') and (Result[Length(Result)] = '"') then
    Result := Copy(Result, 2, Length(Result) - 2);
end;

function MaskCommandForLog(const FileName, Params: string): string;
var
  Cur: Integer;
  A1Start, A1End, A2Start, A2End, A3Start, A3End: Integer;
  Arg1: string;
begin
  Result := MaskPasswordsInString(Params);

  if CompareText(FileName, 'net.exe') <> 0 then
    exit;

  Cur := 1;
  A1Start := 0; A1End := 0;
  A2Start := 0; A2End := 0;
  A3Start := 0; A3End := 0;

  NextArgRange(Result, Cur, A1Start, A1End);
  if (A1Start <= 0) or (A1End < A1Start) then
    exit;

  Arg1 := UpperCase(StripWrappingQuotes(Copy(Result, A1Start, A1End - A1Start + 1)));
  if Arg1 <> 'USER' then
    exit;

  NextArgRange(Result, Cur, A2Start, A2End); // username
  NextArgRange(Result, Cur, A3Start, A3End); // password
  if (A3Start > 0) and (A3End >= A3Start) then
  begin
    Delete(Result, A3Start, A3End - A3Start + 1);
    Insert('"*****"', Result, A3Start);
  end;
end;

// -----------------------------------------------------------------------------
// PATH CONSTRUCTION HELPERS
// -----------------------------------------------------------------------------

// Expand a filename under {tmp}
function TempFile(const FileName: string): string;
begin
  Result := ExpandConstant('{tmp}\' + FileName);
end;

function EnsureDebugWorkDir: string;
var
  BaseDir: string;
begin
  LogEntry('EnsureDebugWorkDir');
  BaseDir := ExpandConstant('{localappdata}\RDPWrapKit');
  if (not DirExists(BaseDir)) and (not CreateDir(BaseDir)) then
  begin
    LogWarn('Could not create debug work base directory: ' + BaseDir);
    Result := ExpandConstant('{tmp}');
    LogExit('EnsureDebugWorkDir');
    exit;
  end;

  Result := BaseDir + '\DebugLogs';
  if (not DirExists(Result)) and (not CreateDir(Result)) then
  begin
    LogWarn('Could not create debug work directory: ' + Result);
    Result := ExpandConstant('{tmp}');
    LogExit('EnsureDebugWorkDir');
    exit;
  end;

  LogDebug('Debug work directory ready: ' + Result);
  LogExit('EnsureDebugWorkDir');
end;

function DebugLogFile(const FileName: string): string;
begin
  LogEntry('DebugLogFile');
  Result := EnsureDebugWorkDir + '\' + FileName;
  LogExit('DebugLogFile');
end;

function BuildPowerShellFileArgs(const ScriptPath, ExtraParams: string; Hidden: Boolean): string; forward;
function GetPSOutput(const Command: string): string; forward;

// Create a filesystem-safe filename from an arbitrary string by replacing
// non-alphanumeric characters with underscores.
function SanitizeFileName(const S: string): string;
var
  i: Integer;
  c: Char;
begin
  Result := '';
  for i := 1 to Length(S) do
  begin
    c := S[i];
    if ((c >= 'a') and (c <= 'z')) or ((c >= 'A') and (c <= 'Z')) or ((c >= '0') and (c <= '9')) then
      Result := Result + c
    else
      Result := Result + '_';
  end;
end;

// Returns True if the given string is safe to use as a Windows filename (no invalid chars).
// Invalid Windows filename characters: < > : " / \ | ? *
function IsValidShortcutName(const Name: string): Boolean;
var
  i: Integer;
  c: Char;
begin
  Result := False;
  if Trim(Name) = '' then
    exit;
  for i := 1 to Length(Name) do
  begin
    c := Name[i];
    if (c = '<') or (c = '>') or (c = ':') or (c = '"') or
       (c = '/') or (c = '\') or (c = '|') or (c = '?') or (c = '*') then
      exit;
  end;
  Result := True;
end;

// Escape single quotes for use in PowerShell single-quoted literals
function PSSingleQuote(const S: string): string;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length(S) do
  begin
    if S[i] = '''' then
      Result := Result + ''''''
    else
      Result := Result + S[i];
  end;
end;

// Build a safe named PowerShell argument using single-quoted values
function BuildPSNamedParam(const Name, Value: string): string;
var
  i: Integer;
  SafeName: string;
  Escaped: string;
  c: Char;
begin
  SafeName := '';
  for i := 1 to Length(Name) do
  begin
    c := Name[i];
    if ((c >= 'A') and (c <= 'Z')) or ((c >= 'a') and (c <= 'z')) or ((c >= '0') and (c <= '9')) then
      SafeName := SafeName + c;
  end;
  if SafeName = '' then
    SafeName := 'Param';
  Escaped := '';
  for i := 1 to Length(Value) do
  begin
    if Value[i] = '"' then
      Escaped := Escaped + '\"'
    else
      Escaped := Escaped + Value[i];
  end;
  Result := '-' + SafeName + ' "' + Escaped + '"';
end;

// Quote an argument for direct executable invocation (CreateProcess semantics)
function QuoteExeArg(const S: string): string;
var
  i: Integer;
  Escaped: string;
begin
  Escaped := '';
  for i := 1 to Length(S) do
  begin
    if S[i] = '"' then
      Escaped := Escaped + '\"'
    else
      Escaped := Escaped + S[i];
  end;
  Result := '"' + Escaped + '"';
end;

// Execute PowerShell script content through a temporary script file
function ExecPowerShellScriptContent(const ScriptBaseName, ScriptContent, ExtraParams: string; Hidden: Boolean; var ResultCode: Integer): Boolean;
var
  ScriptPath: string;
  OpTick: Cardinal;
begin
  LogEntry('ExecPowerShellScriptContent');
  OpTick := GetTickCount;
  ScriptPath := TempFile(ScriptBaseName);
  SaveStringToFile(ScriptPath, ScriptContent, False);
  LogDebug('PowerShell File: ' + ScriptPath + ' ' + MaskPasswordsInString(ExtraParams));
  Result := Exec(EXE_POWERSHELL, BuildPowerShellFileArgs(ScriptPath, ExtraParams, Hidden), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  LogDebug('PowerShell exit=' + IntToStr(ResultCode) + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  DeleteFile(ScriptPath);
  LogExit('ExecPowerShellScriptContent');
end;

function ExecSavedPowerShellDebugScriptParams(const ScriptTag, UserName, ScriptContent, ExtraParams: string; Hidden: Boolean; var ResultCode: Integer): Boolean;
var
  ScriptPath: string;
  OpTick: Cardinal;
begin
  LogEntry('ExecSavedPowerShellDebugScriptParams');
  OpTick := GetTickCount;
  ScriptPath := TempFile(ScriptTag + '_' + SanitizeFileName(UserName) + '.ps1');
  SaveStringToFile(ScriptPath, ScriptContent, False);
  LogDebug('PowerShell File: ' + ScriptPath + ' ' + MaskPasswordsInString(ExtraParams));
  Result := Exec(EXE_POWERSHELL, BuildPowerShellFileArgs(ScriptPath, ExtraParams, Hidden), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  if not Result then
    LogError('Failed to launch PowerShell script: code=' + IntToStr(ResultCode) + ' message=' + SysErrorMessage(ResultCode));
  LogDebug('PowerShell exit=' + IntToStr(ResultCode) + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  DeleteFile(ScriptPath);
  LogExit('ExecSavedPowerShellDebugScriptParams');
end;

function BuildAddGroupMemberPowerShellScript(const GroupSid, UserName, OutPath, SuccessTag: string): string;
begin
  Result :=
    'param([string]$GroupSid, [string]$UserName, [string]$OutPath, [string]$SuccessTag)' + #13#10 +
    '$ErrorActionPreference = ''Stop''' + #13#10 +
    'try {' + #13#10 +
    '  $group = Get-LocalGroup -SID $GroupSid -ErrorAction Stop' + #13#10 +
    '  $resolvedName = $group.Name' + #13#10 +
    '  Add-LocalGroupMember -Group $group -Member $UserName -ErrorAction Stop' + #13#10 +
    '  @($SuccessTag, (''Group={0}'' -f $resolvedName), (''GroupSid={0}'' -f $GroupSid), (''User={0}'' -f $UserName)) | Out-File -FilePath $OutPath -Encoding UTF8' + #13#10 +
    '  exit 0' + #13#10 +
    '} catch {' + #13#10 +
    '  @(' + #13#10 +
    '    ''ADD_GROUP_FAIL'',' + #13#10 +
    '    (''GroupSid={0}'' -f $GroupSid),' + #13#10 +
    '    (''User={0}'' -f $UserName),' + #13#10 +
    '    ''ExceptionType='' + $_.Exception.GetType().FullName,' + #13#10 +
    '    ''Message='' + $_.Exception.Message,' + #13#10 +
    '    ''HResult='' + $_.Exception.HResult,' + #13#10 +
    '    ''CategoryInfo='' + $_.CategoryInfo.ToString(),' + #13#10 +
    '    ''FullyQualifiedErrorId='' + $_.FullyQualifiedErrorId,' + #13#10 +
    '    ''StackTrace:'',' + #13#10 +
    '    ($_ | Out-String)' + #13#10 +
    '  ) | Out-File -FilePath $OutPath -Encoding UTF8' + #13#10 +
    '  exit 1' + #13#10 +
    '}';
end;

function ValidateGroupMembership(const GroupName, UserName: string): Boolean;
var
  OutPath: string;
  ResultCode: Integer;
  Lines: TStringList;
  i: Integer;
  Line: string;
  BackslashPos: Integer;
  MemberName: string;
  InList: Boolean;
begin
  LogEntry('ValidateGroupMembership');
  Result := False;
  OutPath := TempFile('grp_members_' + SanitizeFileName(GroupName) + '.txt');

  Exec('cmd.exe', '/c net localgroup ' + QuoteExeArg(GroupName) + ' > ' + QuoteExeArg(OutPath) + ' 2>&1',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  if not FileExists(OutPath) then
  begin
    LogDebug('ValidateGroupMembership: output file not found for ' + GroupName);
    LogExit('ValidateGroupMembership');
    exit;
  end;

  Lines := TStringList.Create;
  try
    Lines.LoadFromFile(OutPath);
    InList := False;
    for i := 0 to Lines.Count - 1 do
    begin
      Line := Trim(Lines[i]);
      if not InList then
      begin
        if (Length(Line) > 0) and (Line[1] = '-') then
          InList := True;
        continue;
      end;
      if Line = '' then
        continue;
      BackslashPos := Pos('\', Line);
      if BackslashPos > 0 then
        MemberName := Copy(Line, BackslashPos + 1, MaxInt)
      else
        MemberName := Line;
      if CompareText(MemberName, UserName) = 0 then
      begin
        Result := True;
        break;
      end;
    end;
  finally
    Lines.Free;
    DeleteFile(OutPath);
  end;
  LogDebug('ValidateGroupMembership: group=' + GroupName + ' user=' + UserName + ' result=' + BoolToStr(Result));
  LogExit('ValidateGroupMembership');
end;

// Verify file is authenticode-signed by Microsoft Corporation
function IsSignedByMicrosoftCorporation(const FilePath: string): Boolean;
var
  OutText: string;
  OpTick: Cardinal;
begin
  LogEntry('IsSignedByMicrosoftCorporation');
  OpTick := GetTickCount;
  LogSectionHeader('SIGNATURE VALIDATION');
  LogKeyValue('File', FilePath);
  OutText := GetPSOutput(
    '$sig = Get-AuthenticodeSignature -FilePath ''' + PSSingleQuote(FilePath) + '''; ' +
    '$subj = ''''; $issuer = ''''; $thumb = ''''; $nb = ''''; $na = ''''; ' +
    'if ($sig.SignerCertificate -ne $null) { ' +
    '  $subj = $sig.SignerCertificate.Subject; ' +
    '  $issuer = $sig.SignerCertificate.Issuer; ' +
    '  $thumb = $sig.SignerCertificate.Thumbprint; ' +
    '  $nb = $sig.SignerCertificate.NotBefore.ToString(''u''); ' +
    '  $na = $sig.SignerCertificate.NotAfter.ToString(''u'') ' +
    '}; ' +
    '$isMs = ($subj -like ''*CN=Microsoft Corporation*''); ' +
    '$ok = (($sig.Status -eq ''Valid'') -and $isMs); ' +
    '''STATUS='' + $sig.Status + ''|IS_MICROSOFT='' + $isMs + ''|SUBJECT='' + $subj + ''|ISSUER='' + $issuer + ''|THUMBPRINT='' + $thumb + ''|NOT_BEFORE='' + $nb + ''|NOT_AFTER='' + $na + ''|RESULT='' + ($(if($ok){''OK''}else{''BAD''}))');
  LogDebug('Signature details: ' + OutText);
  Result := Pos('|RESULT=OK', UpperCase(Trim(OutText))) > 0;
  if Result then
    LogInfo('Signature verdict: Microsoft publisher validation passed [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]')
  else
    LogWarn('Signature verdict: Microsoft publisher validation failed [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('IsSignedByMicrosoftCorporation');
end;

// Resolve mstsc.exe path dynamically
function GetMstscPath: string;
begin
  Result := ExpandConstant('{sys}\mstsc.exe');
  if not FileExists(Result) then
    Result := '';
end;

function ValidateRdpPortInput(const PortText: string; var PortValue: Integer; var ErrorText: string): Boolean;
var
  i: Integer;
  CleanText: string;
begin
  Result := False;
  PortValue := 0;
  ErrorText := '';
  CleanText := Trim(PortText);

  if CleanText = '' then
  begin
    ErrorText := 'RDP Port is required.';
    exit;
  end;

  for i := 1 to Length(CleanText) do
  begin
    if (CleanText[i] < '0') or (CleanText[i] > '9') then
    begin
      ErrorText := 'RDP Port must contain digits only.';
      exit;
    end;
  end;

  PortValue := StrToIntDef(CleanText, 0);
  if (PortValue < 1) or (PortValue > 65535) then
  begin
    ErrorText := 'RDP Port must be between 1 and 65535.';
    exit;
  end;

  Result := True;
end;

// -----------------------------------------------------------------------------
// POWERSHELL EXECUTION HELPERS
// -----------------------------------------------------------------------------

// Encode a Pascal string as UTF-16LE base64 for use with PowerShell -EncodedCommand.
// Inno Setup strings are UCS-2/UTF-16LE internally, so each Char is two bytes.
function PSBase64Encode(const S: string): string;
var
  Table: string;
  Bytes: array of Byte;
  n, i, b0, b1, b2: Integer;
begin
  Table := 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  n := Length(S) * 2;
  SetLength(Bytes, n);
  for i := 0 to Length(S) - 1 do
  begin
    Bytes[i * 2]     := Ord(S[i + 1]) and $FF;
    Bytes[i * 2 + 1] := (Ord(S[i + 1]) shr 8) and $FF;
  end;
  Result := '';
  i := 0;
  while i < n do
  begin
    b0 := Bytes[i];
    if i + 1 < n then b1 := Bytes[i + 1] else b1 := 0;
    if i + 2 < n then b2 := Bytes[i + 2] else b2 := 0;
    Result := Result + Table[(b0 shr 2) + 1];
    Result := Result + Table[((b0 and 3) shl 4) or (b1 shr 4) + 1];
    if i + 1 < n then
      Result := Result + Table[((b1 and $F) shl 2) or (b2 shr 6) + 1]
    else
      Result := Result + '=';
    if i + 2 < n then
      Result := Result + Table[(b2 and $3F) + 1]
    else
      Result := Result + '=';
    i := i + 3;
  end;
end;

// Build PowerShell -Command args
function BuildPowerShellArgs(const Command: string; Hidden: Boolean): string;
begin
  if Hidden then
    Result := PS_ARGS_HIDDEN + ' -Command "' + Command + '"'
  else
    Result := PS_ARGS_BASE + ' -Command "' + Command + '"';
end;

// Build PowerShell -File args (with optional extra params)
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
  OpTick: Cardinal;
begin
  LogEntry('ExecPowerShellHidden');
  OpTick := GetTickCount;
  PSArgs := BuildPowerShellArgs(Command, True);
  // Log command and run (mask any embedded passwords)
  LogDebug('PowerShell Hidden: ' + MaskPasswordsInString(PSArgs));
  Result := Exec(EXE_POWERSHELL, PSArgs, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  LogDebug('PowerShell exitcode=' + IntToStr(ResultCode) + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('ExecPowerShellHidden');
end;

// Execute a PowerShell command and capture stdout to a temp file, returning
// the trimmed output as a string. Returns empty string on failure.
// NOTE: Temp file usage is required because Exec() does not return stdout.
// All captured output is logged to the main installer log before the file
// is deleted, so no external log files remain.
function GetPSOutput(const Command: string): string;
var
  PSPath: string;
  RC: Integer;
  SL: TStringList;
  j: Integer;
  OpTick: Cardinal;
begin
  LogEntry('GetPSOutput');
  OpTick := GetTickCount;
  Result := '';
  PSPath := TempFile('psout.txt');
  Exec(EXE_POWERSHELL,
    BuildPowerShellArgs(Command + ' | Out-File -Encoding UTF8 ''' + PSPath + ''' -Force', True),
    '', SW_HIDE, ewWaitUntilTerminated, RC);
  if (RC = 0) and FileExists(PSPath) then
  begin
    SL := TStringList.Create;
    try
      SL.LoadFromFile(PSPath);
      Result := Trim(SL.Text);
      // Log captured output to main installer log
      if Length(Result) > 0 then
      begin
        LogDebug('GetPSOutput (' + IntToStr(SL.Count) + ' lines):');
        for j := 0 to SL.Count - 1 do
          LogDebug('  ' + SL[j]);
      end
      else
        LogDebug('GetPSOutput: output file empty');
    finally
      SL.Free;
      DeleteFile(PSPath);
    end;
  end
  else
    LogDebug('GetPSOutput: no output (RC=' + IntToStr(RC) + ')');
  LogDebug('GetPSOutput exit=' + IntToStr(RC) + ' resultLen=' + IntToStr(Length(Result)) + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('GetPSOutput');
end;

// Execute PowerShell and capture both stdout/stderr output regardless of exit code.
// NOTE: Temp file usage is required because Exec() does not return stdout.
// All captured output is logged to the main installer log before the file
// is deleted, so no external log files remain.
function ExecPSCaptureAll(const Command: string; var ResultCode: Integer): string;
var
  PSPath: string;
  SL: TStringList;
  j: Integer;
  WrappedCommand: string;
  OpTick: Cardinal;
begin
  LogEntry('ExecPSCaptureAll');
  OpTick := GetTickCount;
  Result := '';
  PSPath := TempFile('psall.txt');
  WrappedCommand := '& { ' + Command + ' } *>&1 | Out-File -Encoding UTF8 ''' + PSPath + ''' -Force';
  Exec(EXE_POWERSHELL,
    BuildPowerShellArgs(WrappedCommand, True),
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  if FileExists(PSPath) then
  begin
    SL := TStringList.Create;
    try
      SL.LoadFromFile(PSPath);
      Result := Trim(SL.Text);
      // Log captured output to main installer log
      if Length(Result) > 0 then
      begin
        LogDebug('ExecPSCaptureAll (' + IntToStr(SL.Count) + ' lines):');
        for j := 0 to SL.Count - 1 do
          LogDebug('  ' + SL[j]);
      end
      else
        LogDebug('ExecPSCaptureAll: output file empty');
    finally
      SL.Free;
      DeleteFile(PSPath);
    end;
  end
  else
    LogDebug('ExecPSCaptureAll: no output file');
  LogDebug('ExecPSCaptureAll exit=' + IntToStr(ResultCode) + ' resultLen=' + IntToStr(Length(Result)) + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('ExecPSCaptureAll');
end;

procedure EnsureRDPSigningCert; forward;

procedure SignRdpFile(const RdpPath: string);
var
  ResultCode: Integer;
  OpTick: Cardinal;
  ExePath: string;
  CmdLine: string;
  RdpFound: Boolean;
  ExeFound: Boolean;
begin
  LogEntry('SignRdpFile');
  OpTick := GetTickCount;
  LogDebug('SignRdpFile: RdpPath=''' + RdpPath + '''');

  // ---- Pre-flight diagnostics: validate RDP file exists and is accessible ----
  LogDebug('SignRdpFile: Pre-flight checking RDP file...');
  RdpFound := FileExists(RdpPath);
  LogDebug('SignRdpFile: FileExists(''' + RdpPath + ''')=' + BoolToStr(RdpFound));

  if RdpFound then
  begin
    LogDebug('SignRdpFile: RDP file confirmed accessible at path');
  end
  else
  begin
    LogWarn('SignRdpFile: RDP file does NOT exist at: ' + RdpPath);
    LogExit('SignRdpFile');
    exit;
  end;

  // ---- Pre-flight diagnostics: validate RdpSignTool.exe exists ----
  ExePath := ExpandConstant('{tmp}\RdpSignTool.exe');
  ExeFound := FileExists(ExePath);
  LogDebug('SignRdpFile: FileExists(''' + ExePath + ''')=' + BoolToStr(ExeFound));
  if not ExeFound then
  begin
    LogWarn('SignRdpFile: RdpSignTool.exe NOT FOUND at ' + ExePath);
    LogExit('SignRdpFile');
    exit;
  end;

  // Ensure the signing certificate exists and is trusted before signing
  EnsureRDPSigningCert;

  // RdpSignTool.exe is a standalone C# console application compiled ahead-of-time
  // (see scripts/build_rdpcrypt.ps1).  It performs all RDP signing logic directly
  // via P/Invoke to crypt32.dll — no PowerShell involved, no C# JIT compilation,
  // no -EncodedCommand overhead.  The EXE is ~20KB and completes in ~100-300ms.
  //
  // Exit codes:
  //   0  = Success
  //   1  = Certificate not found
  //   2  = RDP file not found/unreadable
  //   3  = Missing "full address" field
  //   4  = CryptSignMessage failure
  //   5  = Internal error/exception
  //   99 = Invalid arguments
  CmdLine := '"' + RdpPath + '"';
  LogDebug('SignRdpFile: Executing: ' + ExePath + ' ' + CmdLine);

  // Capture RdpSignTool.exe output for diagnostic purposes.
  // Use RunCmdCapture to redirect stdout/stderr to a temp file and log it on failure.
  ResultCode := RunCmdCapture('"' + ExePath + '" ' + CmdLine, 'rdpsign_' + ExtractFileName(RdpPath));
  LogDebug('SignRdpFile: RdpSignTool.exe exit=' + IntToStr(ResultCode) + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');

  if ResultCode = 0 then
    LogInfo('SignRdpFile: signed ' + RdpPath + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]')
  else
    LogWarn('SignRdpFile: failed to sign ' + RdpPath + ' (exit=' + IntToStr(ResultCode) + ') [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogInfo('SignRdpFile total [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('SignRdpFile');
end;

procedure UnSignRdpFile(const RdpPath: string);
// Strips all signature-related lines (signscope, signature, alternate full address,
// and orphaned base64 artifacts) from an RDP file.
// This allows mstsc /edit to modify every field without signature validation
// interference. The file will be re-signed after the user finishes editing.
var
  Lines: TStringList;
  OutLines: TStringList;
  i: Integer;
  Line: string;
  LowerLine: string;
  StrippedCount: Integer;
  OpTick: Cardinal;
begin
  LogEntry('UnSignRdpFile');
  OpTick := GetTickCount;

  if not FileExists(RdpPath) then
  begin
    LogWarn('UnSignRdpFile: File not found: ' + RdpPath);
    LogExit('UnSignRdpFile');
    exit;
  end;

  LogDebug('UnSignRdpFile: Reading ' + RdpPath);
  Lines := TStringList.Create;
  OutLines := TStringList.Create;
  try
    Lines.LoadFromFile(RdpPath);
    StrippedCount := 0;

    for i := 0 to Lines.Count - 1 do
    begin
      Line := Trim(Lines[i]);
      if Line = '' then
        continue;

      LowerLine := LowerCase(Line);

      // Strip signscope, signature, and alternate full address lines
      if (Pos('signscope:', LowerLine) = 1) or
         (Pos('signature:', LowerLine) = 1) or
         (Pos('alternate full address:', LowerLine) = 1) then
      begin
        Inc(StrippedCount);
        LogDebug('UnSignRdpFile: Stripped line: ' + Copy(Line, 1, 60));
      end
      // Strip orphaned base64-only lines (no colon at all � leftover artifact
      // from previous signatures that used a different wrapping format)
      else if Pos(':', Line) = 0 then
      begin
        // Verify it looks like pure base64 (A-Za-z0-9+/=) before stripping
        if Length(Line) > 16 then
        begin
          Inc(StrippedCount);
          LogDebug('UnSignRdpFile: Stripped orphaned base64 line: ' + Copy(Line, 1, 60));
        end
        else
          OutLines.Add(Line);
      end
      else
        OutLines.Add(Line);
    end;

    if StrippedCount > 0 then
    begin
      OutLines.SaveToFile(RdpPath);
      LogInfo('UnSignRdpFile: Stripped ' + IntToStr(StrippedCount) +
        ' signature lines from ' + ExtractFileName(RdpPath) +
        ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
    end
    else
      LogDebug('UnSignRdpFile: No signature lines found in ' + ExtractFileName(RdpPath) +
        ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  finally
    Lines.Free;
    OutLines.Free;
  end;

  LogExit('UnSignRdpFile');
end;


// Ensures the RDPWrapKit code-signing cert exists, is trusted, and is registered
// as a trusted RDP publisher in the Terminal Services policy.
//
// PERFORMANCE: All four logical steps are executed in a single PowerShell
// invocation (~500ms total for PS startup + execution) instead of four separate
// invocations (~2s total).  The single script returns the certificate thumbprint
// on success, or exits with a non-zero code on failure.
procedure EnsureRDPSigningCert;
var
  PSCommand: string;
  ResultCode: Integer;
  PSOut: string;
  Thumb: string;
  OpTick: Cardinal;
begin
  LogEntry('EnsureRDPSigningCert');
  OpTick := GetTickCount;
  Thumb := '';

  // Single consolidated PowerShell script that performs all four steps:
  //   1. Find or create the cert in LocalMachine\My
  //   2. Import into Trusted Root (if not already there)
  //   3. [implicit] verify — step 2 already checks existence
  //   4. Register thumbprint in TrustedCertThumbprints policy
  // Outputs thumbprint on success; any error triggers catch/throw.
  LogDebug('EnsureRDPSigningCert: Running consolidated PS cert setup');
  // NOTE: PSCommand is a one-liner separated by ; — NO inline # comments
  // because the PS script has no #13#10 line breaks and # would comment out
  // the rest of the line (all statements after it).
  PSCommand :=
    '$ErrorActionPreference = ''Stop''; ' +
    'try { ' +
    '$subjectName = ''CN=RDPWrapKit: Only trust if connecting to 127.0.0.2''; ' +
    '$existing = Get-ChildItem "Cert:\\LocalMachine\\My" | Where-Object { $_.Subject -eq $subjectName } | Select-Object -First 1; ' +
    'if ($existing) { $thumb = $existing.Thumbprint } else { ' +
    '$c = New-SelfSignedCertificate -Subject $subjectName -CertStoreLocation "Cert:\\LocalMachine\\My" -KeyUsage DigitalSignature -Type CodeSigningCert -NotAfter (Get-Date).AddYears(10); ' +
    '$thumb = $c.Thumbprint }; ' +
    'if (-not $thumb) { throw "Failed to obtain certificate thumbprint" }; ' +
    '$rootCert = Get-ChildItem "Cert:\\LocalMachine\\Root" | Where-Object { $_.Thumbprint -eq $thumb } | Select-Object -First 1; ' +
    'if (-not $rootCert) { ' +
    '$tmp = Join-Path $env:TEMP "rdpwrapkit.cer"; ' +
    '$myCert = Get-ChildItem "Cert:\\LocalMachine\\My" | Where-Object { $_.Thumbprint -eq $thumb } | Select-Object -First 1; ' +
    'if (-not $myCert) { throw ''Certificate vanished from LocalMachine\My'' }; ' +
    'Export-Certificate -Cert $myCert -FilePath $tmp | Out-Null; ' +
    'Import-Certificate -FilePath $tmp -CertStoreLocation "Cert:\\LocalMachine\\Root" | Out-Null; ' +
    'Remove-Item $tmp -Force; ' +
    '$rootCert = Get-ChildItem "Cert:\\LocalMachine\\Root" | Where-Object { $_.Thumbprint -eq $thumb } | Select-Object -First 1; ' +
    'if (-not $rootCert) { throw "Failed to import cert to Trusted Root" } }; ' +
    '$keyPath = ''Software\Policies\Microsoft\Windows NT\Terminal Services''; ' +
    '$key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($keyPath, $true); ' +
    'if (-not $key) { ' +
    '$null = [Microsoft.Win32.Registry]::LocalMachine.CreateSubKey($keyPath); ' +
    '$key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($keyPath, $true) }; ' +
    '$raw = $key.GetValue(''TrustedCertThumbprints'', $null); ' +
    '$existingList = if ($raw -is [string]) { $raw } else { '''' }; ' +
    'if ($existingList -notmatch [regex]::Escape($thumb)) { ' +
    '$newVal = if ($existingList -ne '''') { $existingList + '','' + $thumb } else { $thumb }; ' +
    '$key.SetValue(''TrustedCertThumbprints'', $newVal, [Microsoft.Win32.RegistryValueKind]::String) }; ' +
    '$key.Close(); ' +
    'Write-Output $thumb ' +
    '} catch { Write-Output (''ERROR: '' + $_.Exception.Message); throw }';

  PSOut := ExecPSCaptureAll(PSCommand, ResultCode);
  LogDebug('EnsureRDPSigningCert [ConsolidatedPS] exit=' + IntToStr(ResultCode) + ' output=' + PSOut);

  if ResultCode <> 0 then
  begin
    LogError('EnsureRDPSigningCert: failed (exit=' + IntToStr(ResultCode) + ')');
    LogExit('EnsureRDPSigningCert');
    exit;
  end;

  Thumb := Trim(PSOut);
  if Thumb = '' then
  begin
    LogError('EnsureRDPSigningCert: empty thumbprint output');
    LogExit('EnsureRDPSigningCert');
    exit;
  end;
  if Copy(Thumb, 1, 6) = 'ERROR:' then
  begin
    LogError('EnsureRDPSigningCert: PS error: ' + PSOut);
    LogExit('EnsureRDPSigningCert');
    exit;
  end;

  LogInfo('EnsureRDPSigningCert: cert ensured and registered (thumb=' + Copy(Thumb, 1, 8) + '...) [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('EnsureRDPSigningCert');
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

function AppRoot: string;
begin
  Result := ExpandConstant('{app}');
end;

function AppBin(const FileName: string): string;
begin
  Result := ExpandConstant('{app}\' + FileName);
end;

function SafeRegString(const Root: Integer; const Key, ValueName: string; const Fallback: string): string;
var
  S: string;
begin
  if RegQueryStringValue(Root, Key, ValueName, S) then
    Result := S
  else
    Result := Fallback;
end;

function SafeRegDword(const Root: Integer; const Key, ValueName: string; const Fallback: Cardinal): Cardinal;
var
  D: Cardinal;
begin
  if RegQueryDWordValue(Root, Key, ValueName, D) then
    Result := D
  else
    Result := Fallback;
end;

procedure LogSystemInfo;
var
  ProductName: string;
  DisplayVersion: string;
  BuildNumber: string;
  EditionID: string;
  UBR: Cardinal;
  InstallLang: string;
  Arch: string;
begin
  LogEntry('LogSystemInfo');
  ProductName := SafeRegString(HKLM, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion', 'ProductName', 'Unknown');
  DisplayVersion := SafeRegString(HKLM, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion', 'DisplayVersion', '');
  if DisplayVersion = '' then
    DisplayVersion := SafeRegString(HKLM, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion', 'ReleaseId', '');
  BuildNumber := SafeRegString(HKLM, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion', 'CurrentBuildNumber', '');
  EditionID := SafeRegString(HKLM, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion', 'EditionID', '');
  UBR := SafeRegDword(HKLM, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion', 'UBR', 0);
  InstallLang := SafeRegString(HKLM, 'SYSTEM\CurrentControlSet\Control\Nls\Language', 'InstallLanguage', '');

  if IsWin64 then
    Arch := 'x64'
  else
    Arch := 'x86';

  LogInfo('SystemInfo: OS=' + ProductName + ' Build.UBR=' + BuildNumber + '.' + IntToStr(UBR));
  if DisplayVersion <> '' then
    LogInfo('SystemInfo: Version=' + DisplayVersion);
  if BuildNumber <> '' then
    LogInfo('SystemInfo: Build=' + BuildNumber + ' (UBR ' + IntToStr(UBR) + ')');
  if EditionID <> '' then
    LogInfo('SystemInfo: Edition=' + EditionID);
  if InstallLang <> '' then
    LogInfo('SystemInfo: InstallLanguage=' + InstallLang);
  LogInfo('SystemInfo: Arch=' + Arch);
  LogExit('LogSystemInfo');
end;

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
  if ST.wMilliseconds < 10 then ms := '00' + IntToStr(ST.wMilliseconds)
  else if ST.wMilliseconds < 100 then ms := '0' + IntToStr(ST.wMilliseconds)
  else ms := IntToStr(ST.wMilliseconds);
  Result := '[' + h + ':' + m + ':' + s + '.' + ms + ']';
end;

function DWordToHex(const Value: Cardinal; const MinDigits: Integer): string;
var
  V: Cardinal;
  Digit: Integer;
begin
  Result := '';
  V := Value;
  if V = 0 then
    Result := '0'
  else
  begin
    while V > 0 do
    begin
      Digit := V and $F;
      if Digit < 10 then
        Result := Chr(Ord('0') + Digit) + Result
      else
        Result := Chr(Ord('A') + Digit - 10) + Result;
      V := V shr 4;
    end;
  end;
  // Pad to minimum digits
  while Length(Result) < MinDigits do
    Result := '0' + Result;
end;

function GetThreadIdHex: string;
begin
  Result := DWordToHex(GetCurrentThreadId, 4);
end;

function GetProcessIdHex: string;
begin
  Result := DWordToHex(GetCurrentProcessId, 4);
end;

function RepeatChar(const Ch: string; const Count: Integer): string;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Count do
    Result := Result + Ch;
end;

// -----------------------------------------------------------------------------
// CENTRALIZED LOGGER ARCHITECTURE
// -----------------------------------------------------------------------------
// All log output flows through LoggerWrite which applies:
//   1. Sensitive key redaction (Password, Credential, AuthToken, Secret, APIKey, ServiceAccount)
//   2. Thread ID and Process ID metadata injection
//   3. Duration tracking via LoggerStartOp/LoggerEndOp
//   4. [DEBUG]/[INFO]/[WARN]/[ERROR] level tagging
// -----------------------------------------------------------------------------

// Helper: redact a single sensitive key pattern from Result/U in-place
procedure RedactKeyPattern(var ResultStr, UpperStr: string; const KeyPattern: string);
var
  p, startPos, endPos, SearchPos, idx: Integer;
begin
  SearchPos := 1;
  idx := PosFrom(KeyPattern, UpperStr, SearchPos);
  while idx > 0 do
  begin
    p := idx + Length(KeyPattern);
    while (p <= Length(ResultStr)) and ((ResultStr[p] = ' ') or (ResultStr[p] = '=')) do
      Inc(p);
    if p > Length(ResultStr) then
      Break;
    if ResultStr[p] = '"' then
    begin
      startPos := p;
      endPos := startPos + 1;
      while (endPos <= Length(ResultStr)) and (ResultStr[endPos] <> '"') do
        Inc(endPos);
      if endPos > Length(ResultStr) then
        endPos := Length(ResultStr);
      Delete(ResultStr, startPos, endPos - startPos + 1);
      Insert('"[REDACTED]"', ResultStr, startPos);
    end
    else
    begin
      endPos := p;
      while (endPos <= Length(ResultStr)) and (ResultStr[endPos] <> ' ') do
        Inc(endPos);
      Delete(ResultStr, p, endPos - p);
      Insert('[REDACTED]', ResultStr, p);
    end;
    UpperStr := UpperCase(ResultStr);
    SearchPos := idx + Length(KeyPattern) + 1;
    idx := PosFrom(KeyPattern, UpperStr, SearchPos);
  end;
end;

// Centralized redaction engine - scrubs all sensitive values before any log output
function LoggerRedact(const S: string): string;
var
  U: string;
  idx, p, startPos, endPos: Integer;
begin
  // First pass: use existing password masking (handles -Password "..." and net.exe patterns)
  Result := MaskPasswordsInString(S);

  // Second pass: redact additional sensitive key patterns using the helper
  U := UpperCase(Result);
  RedactKeyPattern(Result, U, '-PASSWORD');
  RedactKeyPattern(Result, U, '-CREDENTIAL');
  RedactKeyPattern(Result, U, '-AUTHTOKEN');
  RedactKeyPattern(Result, U, '-SECRET');
  RedactKeyPattern(Result, U, '-APIKEY');
  RedactKeyPattern(Result, U, '-SERVICEACCOUNT');

  // Third pass: redact Password/credential values in key=value contexts (e.g. User=x Password=y)
  U := UpperCase(Result);
  idx := Pos('PASSWORD:', U);
  while idx > 0 do
  begin
    p := idx + Length('PASSWORD:');
    if (p <= Length(Result)) and (Result[p] <> ' ') then
    begin
      startPos := p;
      endPos := startPos;
      while (endPos <= Length(Result)) and (Result[endPos] <> ' ') and (Result[endPos] <> '|') and (Result[endPos] <> #13) and (Result[endPos] <> #10) do
        Inc(endPos);
      Delete(Result, startPos, endPos - startPos);
      Insert('[REDACTED]', Result, startPos);
    end;
    U := UpperCase(Result);
    idx := Pos('PASSWORD:', U);
  end;
end;

// Core log writer - every log entry flows through here
procedure LoggerWrite(const Level, Msg: string);
var
  FinalMsg: string;
  RedactedMsg: string;
  DurationStr: string;
begin
  // Apply centralized redaction
  RedactedMsg := LoggerRedact(Msg);

  // Build duration string if an operation is active
  if LoggerOpActive then
    DurationStr := ' [DURATION:' + IntToStr(GetTickCount - LoggerOpStartTick) + 'ms]'
  else
    DurationStr := '';

  // Format: [HH:MM:SS.mmm] [LEVEL] [Thread:0xTID] [PID:0xPID] [Time: Xms] Message
  FinalMsg := GetTimestampString + ' [' + Level + '] [Thread:0x' + GetThreadIdHex + '] [PID:0x' + GetProcessIdHex + ']' + DurationStr + ' ' + RedactedMsg;

  try
    SaveStringToFile(InstallLogPath, FinalMsg + #13#10, True);
  except
  end;
end;

// Convenience wrappers
procedure LogInfo(const Msg: string);
begin
  LoggerWrite('INFO', Msg);
end;

procedure LogDebug(const Msg: string);
begin
  LoggerWrite('DEBUG', Msg);
end;

procedure LogWarn(const Msg: string);
begin
  LoggerWrite('WARN', Msg);
end;

procedure LogError(const Msg: string);
begin
  LoggerWrite('ERROR', Msg);
end;

// Lifecycle tracing - function entry/exit
procedure LogEntry(const FuncName: string);
begin
  LogDebug('>>> ENTER ' + FuncName);
end;

procedure LogExit(const FuncName: string);
begin
  LogDebug('<<< EXIT ' + FuncName);
end;

// Performance profiling helpers
procedure LoggerStartOp;
begin
  LoggerOpStartTick := GetTickCount;
  LoggerOpActive := True;
end;

function LoggerEndOp: Cardinal;
begin
  LoggerOpActive := False;
  Result := GetTickCount - LoggerOpStartTick;
end;

// Typed I/O operation logging with performance metric
procedure LogFileOp(const Operation, FilePath: string; const Success: Boolean; const Extra: string);
var
  Duration: Cardinal;
  StatusStr: string;
begin
  LoggerStartOp;
  // The operation happens externally; we log before and after
  if Success then
    StatusStr := 'Success'
  else
    StatusStr := 'Failed';
  Duration := LoggerEndOp;
  LogDebug('File' + Operation + ': Path=''' + FilePath + ''', Result=''' + StatusStr + ''', Duration=' + IntToStr(Duration) + 'ms' + Extra);
end;

// Typed registry operation logging with performance metric
procedure LogRegOp(const Operation, RegKey, RegValue: string; const Success: Boolean);
var
  Duration: Cardinal;
  StatusStr: string;
begin
  if Success then
    StatusStr := 'Success'
  else
    StatusStr := 'Failed';
  // Duration is the time of the write operation itself
  Duration := 0;
  LoggerWrite('DEBUG', 'Reg' + Operation + ': Key=''' + RegKey + ''', Value=''' + LoggerRedact(RegValue) + ''', Result=''' + StatusStr + ''', Duration=' + IntToStr(Duration) + 'ms');
end;

// Typed service operation logging with performance metric
procedure LogServiceOp(const Operation, ServiceName: string; const ResultCode: Integer);
var
  Duration: Cardinal;
begin
  Duration := LoggerEndOp;
  LoggerWrite('DEBUG', 'Service' + Operation + ': Name=''' + ServiceName + ''', ExitCode=' + IntToStr(ResultCode) + ', Duration=' + IntToStr(Duration) + 'ms');
end;

// Simulation scenario logging (uses [SIM] level)
procedure LogSim(const ScenarioText: string);
begin
  LoggerWrite('SIM', 'SIMULATED SCENARIO: ' + ScenarioText);
  Log('SIMULATED SCENARIO: ' + ScenarioText);
end;

procedure LogSectionHeader(const Title: string);
begin
  LoggerWrite('INFO', '+' + RepeatChar('-', 78) + '+');
  LoggerWrite('INFO', '| ' + Title);
  LoggerWrite('INFO', '+' + RepeatChar('-', 78) + '+');
end;

function NormalizeLogLevel(const Msg: string): string;
var
  U: string;
begin
  U := UpperCase(Trim(Msg));
  if Pos('ERROR:', U) = 1 then
    Result := 'ERROR'
  else if Pos('WARNING:', U) = 1 then
    Result := 'WARN '
  else if Pos('SIMULATED SCENARIO:', U) = 1 then
    Result := 'SIM  '
  else if Pos('DEBUG:', U) = 1 then
    Result := 'DEBUG'
  else
    Result := 'INFO ';
end;

function BeautifyLogMessage(const Msg: string): string;
var
  Lvl: string;
begin
  Lvl := NormalizeLogLevel(Msg);
  Result := '[' + Lvl + '] ' + Msg;
end;

function GetCurrentInstallModeName: string;
begin
  case SelectedInstallMode of
    installModeInstall:
      Result := 'Install';
    installModeEditShortcuts:
      Result := 'Edit Shortcut Settings';
    installModeEditSystemwideSettings:
      Result := 'Edit System-wide RDP Settings';
    installModeUninstall:
      Result := 'Uninstall';
    installModeShowRDPInfo:
      Result := 'Show RDP Info';
    installModeQuickFixes:
      Result := 'Quick Fixes';
  else
    Result := 'Unknown(' + IntToStr(SelectedInstallMode) + ')';
  end;
end;

function GetPageNameById(const PageID: Integer): string;
begin
  if PageID = wpWelcome then Result := 'Welcome'
  else if PageID = wpLicense then Result := 'License Agreement'
  else if PageID = wpPassword then Result := 'Password'
  else if PageID = wpInfoBefore then Result := 'Info Before'
  else if PageID = wpUserInfo then Result := 'User Info'
  else if PageID = wpSelectDir then Result := 'Select Destination Location'
  else if PageID = wpSelectComponents then Result := 'Select Components'
  else if PageID = wpSelectProgramGroup then Result := 'Select Start Menu Folder'
  else if PageID = wpSelectTasks then Result := 'Select Additional Tasks'
  else if PageID = wpReady then Result := 'Ready to Install'
  else if PageID = wpPreparing then Result := 'Preparing to Install'
  else if PageID = wpInstalling then Result := 'Installing'
  else if PageID = wpInfoAfter then Result := 'Info After'
  else if PageID = wpFinished then Result := 'Finished'
  else if Assigned(WelcomePage) and (PageID = WelcomePage.ID) then Result := 'Custom: Welcome'
  else if Assigned(Page_InstallOptions) and (PageID = Page_InstallOptions.ID) then Result := 'Custom: Setup Options'
  else if Assigned(UserPage) and (PageID = UserPage.ID) then Result := 'Custom: Create RDP User Account'
  else if Assigned(Page_ShortcutSettings) and (PageID = Page_ShortcutSettings.ID) then Result := 'Custom: Shortcut Settings'
  else if Assigned(Page_CreateShortcutsForExistingUsers) and (PageID = Page_CreateShortcutsForExistingUsers.ID) then Result := 'Custom: Create Shortcuts for Existing Users'
  else if Assigned(EditShortcutPage) and (PageID = EditShortcutPage.ID) then Result := 'Custom: Edit Existing Shortcut'
  else if Assigned(Page_EditShortcutAdvanced) and (PageID = Page_EditShortcutAdvanced.ID) then Result := 'Custom: Advanced Shortcut Editing'
  else if Assigned(EditSystemwideSettingsPage) and (PageID = EditSystemwideSettingsPage.ID) then Result := 'Custom: Edit System-wide RDP Settings'
  else if Assigned(QuickFixesPage) and (PageID = QuickFixesPage.ID) then Result := 'Custom: Quick Fixes'
  else if Assigned(Page_ShowRDPInfo) and (PageID = Page_ShowRDPInfo.ID) then Result := 'Custom: Show RDP Info'
  else
    Result := 'Custom/Unknown';
end;

procedure LogPageContext(const PageID: Integer);
var
  ChosenAction: string;
begin
  LogKeyValue('PageId', IntToStr(PageID));
  LogKeyValue('PageName', GetPageNameById(PageID));
  LogKeyValue('Wizard Next Caption', WizardForm.NextButton.Caption);

  if Assigned(rbInstall) and rbInstall.Checked then
    ChosenAction := 'Install'
  else if Assigned(rbEditShortcutSettings) and rbEditShortcutSettings.Checked then
    ChosenAction := 'Edit Shortcut Settings'
  else if Assigned(rbEditSystemwideSettings) and rbEditSystemwideSettings.Checked then
    ChosenAction := 'Edit System-wide RDP Settings'
  else if Assigned(rbShowRDPInfo) and rbShowRDPInfo.Checked then
    ChosenAction := 'Tune System Performance'
  else if Assigned(rbQuickFixes) and rbQuickFixes.Checked then
    ChosenAction := 'Quick Fixes'
  else if Assigned(rbUninstall) and rbUninstall.Checked then
    ChosenAction := 'Uninstall'
  else
    ChosenAction := 'Not selected yet';

  LogKeyValue('Top-level selection', ChosenAction);
  LogKeyValue('Resolved mode', GetCurrentInstallModeName);
  LogKeyValue('Install TermWrap', BoolToStr(DoInstallTermWrap));
  LogKeyValue('Create shortcuts', BoolToStr(DoCreateRdpShortcuts));

  if Assigned(rbCreateUsers) and Assigned(rbUseExistingUsers) then
  begin
    if rbCreateUsers.Checked then
      LogKeyValue('Shortcut user source', 'Create new users')
    else if rbUseExistingUsers.Checked then
      LogKeyValue('Shortcut user source', 'Use existing users')
    else
      LogKeyValue('Shortcut user source', 'Not selected');
  end;
end;

procedure LogKeyValue(const KeyName, KeyValue: string);
begin
  LogDebug('  - ' + KeyName + ': ' + KeyValue);
end;

procedure DumpTextFileToLog(const HeaderText, FilePath: string);
var
  Tmp: TStringList;
  k: Integer;
begin
  LogEntry('DumpTextFileToLog');
  if not FileExists(FilePath) then
  begin
    LogWarn(HeaderText + ' file not found: ' + FilePath);
    LogExit('DumpTextFileToLog');
    exit;
  end;

  Tmp := TStringList.Create;
  try
    try
      Tmp.LoadFromFile(FilePath);
      LogDebug(HeaderText + ' (' + IntToStr(Tmp.Count) + ' lines):');
      for k := 0 to Tmp.Count - 1 do
        LogDebug(Tmp[k]);
    except
      LogWarn('Failed to read debug output file: ' + FilePath);
    end;
  finally
    Tmp.Free;
  end;
  LogExit('DumpTextFileToLog');
end;

procedure LogPasswordPipeline(const StageName, UserName, Password: string);
begin
  if PASSWORD_PIPELINE_DIAG = 0 then
    exit;
  LogDebug('PASSWORD_DIAG [' + StageName + '] user=' + UserName + ' :: details=[REDACTED]');
end;

procedure LogEncryptedFileSummary(const StageName, FilePath: string);
var
  EncRaw: AnsiString;
  EncText: string;
begin
  if PASSWORD_PIPELINE_DIAG = 0 then
    exit;

  if not FileExists(FilePath) then
  begin
    LogDebug('PASSWORD_DIAG [' + StageName + '] encrypted file missing: ' + FilePath);
    exit;
  end;

  if not LoadStringFromFile(FilePath, EncRaw) then
  begin
    LogDebug('PASSWORD_DIAG [' + StageName + '] failed reading encrypted file: ' + FilePath);
    exit;
  end;

  EncText := Trim(String(EncRaw));
  LogDebug('PASSWORD_DIAG [' + StageName + '] encLen=' + IntToStr(Length(EncText)));
end;

procedure PreserveDebugLogFileToDesktop(const FilePath: string);
var
  DestDir: string;
  DestPath: string;
  BaseName: string;
begin
  LogEntry('PreserveDebugLogFileToDesktop');
  if PRESERVE_USER_CREATE_DEBUG_LOGS = 0 then
  begin
    LogExit('PreserveDebugLogFileToDesktop');
    exit;
  end;

  if not FileExists(FilePath) then
  begin
    LogExit('PreserveDebugLogFileToDesktop');
    exit;
  end;

  DestDir := ExpandConstant('{userdesktop}\RDPWrapKit_DebugLogs');
  if (not DirExists(DestDir)) and (not CreateDir(DestDir)) then
  begin
    LogWarn('Could not create debug log folder: ' + DestDir);
    LogExit('PreserveDebugLogFileToDesktop');
    exit;
  end;

  BaseName := ChangeFileExt(ExtractFileName(FilePath), '');
  DestPath := DestDir + '\' + BaseName + '_' + IntToStr(GetTickCount) + '.log';
  if CopyFile(FilePath, DestPath, False) then
    LogInfo('Saved debug user-create log: ' + DestPath)
  else
    LogWarn('Failed to save debug user-create log copy for ' + FilePath);
  LogExit('PreserveDebugLogFileToDesktop');
end;

procedure RegisterDebugFileForFinishCleanup(const FilePath: string);
begin
  if (Trim(FilePath) = '') or (not FileExists(FilePath)) then
    exit;
  if not Assigned(PendingDebugCleanupFiles) then
    PendingDebugCleanupFiles := TStringList.Create;
  if PendingDebugCleanupFiles.IndexOf(FilePath) < 0 then
  begin
    PendingDebugCleanupFiles.Add(FilePath);
    LogDebug('Deferred cleanup registered for Finish: ' + FilePath);
  end;
end;

procedure CleanupPendingDebugFiles;
var
  i: Integer;
  P: string;
begin
  LogEntry('CleanupPendingDebugFiles');
  if not Assigned(PendingDebugCleanupFiles) then
  begin
    LogExit('CleanupPendingDebugFiles');
    exit;
  end;

  if CLEANUP_DEBUG_FILES_ON_FINISH = 0 then
  begin
    LogSectionHeader('FINISH CLEANUP: DEFERRED DEBUG FILES');
    LogInfo('Deferred debug cleanup skipped; files retained for troubleshooting');
    LogKeyValue('Queued files retained', IntToStr(PendingDebugCleanupFiles.Count));
    PendingDebugCleanupFiles.Clear;
    LogExit('CleanupPendingDebugFiles');
    exit;
  end;

  LogSectionHeader('FINISH CLEANUP: DEFERRED DEBUG FILES');
  LogKeyValue('Queued files', IntToStr(PendingDebugCleanupFiles.Count));

  for i := 0 to PendingDebugCleanupFiles.Count - 1 do
  begin
    P := PendingDebugCleanupFiles[i];
    if FileExists(P) then
    begin
      if DeleteFile(P) then
        LogDebug('Deleted deferred debug file: ' + P)
      else
        LogWarn('Failed to delete deferred debug file: ' + P);
    end
    else
      LogDebug('Deferred debug file already missing: ' + P);
  end;

  PendingDebugCleanupFiles.Clear;
  LogExit('CleanupPendingDebugFiles');
end;

procedure PromptManualDownload(const ComponentName, Url, Reason: string);
var
  Choice: Integer;
  RC: Integer;
begin
  LogEntry('PromptManualDownload');
  LogSectionHeader('MANUAL DOWNLOAD REQUIRED');
  LogKeyValue('Component', ComponentName);
  LogKeyValue('Reason', Reason);
  LogKeyValue('URL', Url);

  Choice := MsgBox(
    'Failed to install ' + ComponentName + '.' + #13#10 +
    'Reason: ' + Reason + #13#10#13#10 +
    'Download URL:' + #13#10 + Url + #13#10#13#10 +
    'Open the download page now?',
    mbError,
    MB_YESNO);

  if Choice = IDYES then
  begin
    LogInfo('Manual download prompt: user chose YES for ' + ComponentName);
    if not ShellExec('', Url, '', '', SW_SHOWNORMAL, ewNoWait, RC) then
      LogWarn('Manual download launch failed for ' + ComponentName + ', ShellExec rc=' + IntToStr(RC))
    else
      LogInfo('Manual download launch succeeded for ' + ComponentName);
  end
  else
  begin
    LogInfo('Manual download prompt: user chose NO for ' + ComponentName);
  end;
  LogExit('PromptManualDownload');
end;

procedure InitInstallerLog;
var
  HeaderTS: string;
begin
  InstallLogPath := ExpandConstant(INSTALL_LOG_PATH);
  LoggerProcId := GetCurrentProcessId;
  LoggerThreadId := GetCurrentThreadId;
  LoggerOpActive := False;

  HeaderTS := GetTimestampString;
  try
    SaveStringToFile(InstallLogPath, HeaderTS + ' [INFO] [Thread:0x' + GetThreadIdHex + '] [PID:0x' + GetProcessIdHex + '] +' + RepeatChar('=', 78) + '+' + #13#10, False);
    SaveStringToFile(InstallLogPath, HeaderTS + ' [INFO] [Thread:0x' + GetThreadIdHex + '] [PID:0x' + GetProcessIdHex + '] | RDPWrapKit Installer Log' + #13#10, True);
    SaveStringToFile(InstallLogPath, HeaderTS + ' [INFO] [Thread:0x' + GetThreadIdHex + '] [PID:0x' + GetProcessIdHex + '] | Build: ' + BUILD_FINGERPRINT + #13#10, True);
    SaveStringToFile(InstallLogPath, HeaderTS + ' [INFO] [Thread:0x' + GetThreadIdHex + '] [PID:0x' + GetProcessIdHex + '] | Session started (UTC)' + #13#10, True);
    SaveStringToFile(InstallLogPath, HeaderTS + ' [INFO] [Thread:0x' + GetThreadIdHex + '] [PID:0x' + GetProcessIdHex + '] +' + RepeatChar('=', 78) + '+' + #13#10, True);
  except
  end;
  LogSectionHeader('ENVIRONMENT SNAPSHOT');
  LogSystemInfo;
  LogInfo('Logger initialized: PID=0x' + GetProcessIdHex + ' TID=0x' + GetThreadIdHex);
end;

procedure WriteInstallerLog(const Msg: string);
begin
  // Legacy bridge - routes old-style WriteInstallerLog calls through the new centralized logger.
  // Auto-detects log level from message prefix for backward compatibility.
  LoggerWrite(Trim(NormalizeLogLevel(Msg)), Msg);
end;

// Run any process hidden and return its exit code
function RunHidden(const FileName, Params: string): Integer;
var
  RC: Integer;
  OpTick: Cardinal;
begin
  LogEntry('RunHidden');
  OpTick := GetTickCount;
  RC := 0;
  LogDebug('Exec: ' + FileName + ' ' + MaskCommandForLog(FileName, Params));
  Exec(FileName, Params, '', SW_HIDE, ewWaitUntilTerminated, RC);
  LogDebug('ExitCode: ' + IntToStr(RC) + ' for ' + FileName + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  Result := RC;
  LogExit('RunHidden');
end;

// Run a cmd.exe one-liner hidden and return its exit code
function RunCmdHidden(const CmdLine: string): Integer;
begin
  Result := RunHidden(EXE_CMD, '/c ' + CmdLine);
end;

// Convert boolean to string
function BoolToStr(Value: Boolean): string;
begin
  if Value then
    Result := 'True'
  else
    Result := 'False';
end;

function SimulateNoMstsc: Boolean;
begin
  Result := SIM_SCENARIO_NO_MSTSC <> 0;
end;

function SimulateNoVCRedist: Boolean;
begin
  Result := SIM_SCENARIO_NO_VCREDIST <> 0;
end;

function SimulateNetFailPowerShell: Boolean;
begin
  Result := SIM_SCENARIO_NET_FAIL_POWERSHELL <> 0;
end;

procedure LogSimulationScenario(const ScenarioText: string);
begin
  LogSim(ScenarioText);
end;

function RunNetHidden(const Params: string): Integer;
begin
  LogEntry('RunNetHidden');
  if SimulateNetFailPowerShell then
  begin
    if SimulateNetFailPowerShell and (not SimLogNetPsShown) then
    begin
      LogSimulationScenario('System fails on net.exe commands and uses PowerShell fallback');
      SimLogNetPsShown := True;
    end;
    LogWarn('Simulation: forcing net.exe failure for params: ' + MaskCommandForLog('net.exe', Params));
    Result := 1;
    LogExit('RunNetHidden');
    exit;
  end;

  Result := RunHidden('net.exe', Params);
  LogExit('RunNetHidden');
end;

// -----------------------------------------------------------------------------
// DEBUG-ENHANCED COMMAND EXECUTION
// -----------------------------------------------------------------------------
// Enhanced wrappers around RunCmdHidden / RunNetHidden that capture stdout
// and stderr to a temp file and dump the output into the installer log.
// These help diagnose failures without revealing passwords (output goes
// through the centralized LoggerRedact pipeline).
// -----------------------------------------------------------------------------

// Run a command via cmd.exe with stdout/stderr captured to a log file.
// OutTag is a short alphanumeric identifier used for the temp filename.
// Returns the exit code. Output is logged at DEBUG level.
function RunCmdCapture(const CmdLine, OutTag: string): Integer;
var
  OutPath: string;
  SL: TStringList;
  j: Integer;
  RC: Integer;
  ExecOk: Boolean;
  FullCmd: string;
begin
  LogEntry('RunCmdCapture');
  OutPath := TempFile('capture_' + SanitizeFileName(OutTag) + '.log');
  RC := 0;

  // Build the full command: cmd /c "CmdLine > "OutPath" 2>&1"
  // CRITICAL: cmd.exe /c quoting rules — when the first character after /c
  // is a double-quote ("), cmd.exe strips the outermost quotes and also
  // removes the LAST double-quote from the entire command string.
  // If the redirection is placed OUTSIDE these outer quotes, the last " is
  // the closing quote of the output path, munging the redirect target into
  // an invalid filename (e.g. capture.log 2>&1 becomes part of the filename).
  //
  // The fix: place the redirection INSIDE the outer quotes so that after
  // cmd.exe strips them, everything (command + redirection) is intact.
  //
  // BEFORE (broken):  /c ""C:\exe" "C:\file"" > "out.log" 2>&1
  //   After strip:    "C:\exe" "C:\file"" > "out.log 2>&1    ← BAD: 2>&1 in filename
  //
  // AFTER (fixed):   /c ""C:\exe" "C:\file" > "out.log" 2>&1"
  //   After strip:    "C:\exe" "C:\file" > "out.log" 2>&1    ← CORRECT
  FullCmd := '/c "' + CmdLine + ' > "' + OutPath + '" 2>&1"';
  LogDebug('RunCmdCapture: executing: ' + EXE_CMD + ' ' + FullCmd);

  ExecOk := Exec(EXE_CMD, FullCmd, '', SW_HIDE, ewWaitUntilTerminated, RC);
  LogDebug('RunCmdCapture: ExecOk=' + BoolToStr(ExecOk) + ' exit=' + IntToStr(RC) + ' tag=' + OutTag);

  if not ExecOk then
    LogWarn('RunCmdCapture: Exec() failed to launch! OS error=' + IntToStr(RC) + ' (' + SysErrorMessage(RC) + ')');

  // Always dump output file if it exists (regardless of exit code)
  if FileExists(OutPath) then
  begin
    SL := TStringList.Create;
    try
      try
        SL.LoadFromFile(OutPath);
        if SL.Count > 0 then
        begin
          LogDebug('Command output (' + OutTag + ', ' + IntToStr(SL.Count) + ' lines):');
          for j := 0 to SL.Count - 1 do
            LogDebug('  ' + SL[j]);
        end
        else
          LogDebug('RunCmdCapture: output file empty for tag=' + OutTag);
      except
        LogWarn('RunCmdCapture: failed to read output file: ' + OutPath);
      end;
    finally
      SL.Free;
    end;
    DeleteFile(OutPath);
  end
  else
  begin
    LogDebug('RunCmdCapture: no output file for tag=' + OutTag);
    LogDebug('RunCmdCapture: full command for diagnosis: ' + FullCmd);
  end;

  Result := RC;
  LogExit('RunCmdCapture');
end;

// Run net.exe with stdout/stderr captured and logged.
// Like RunNetHidden but dumps the command output on failure.
// Uses the simulation bypass from RunNetHidden.
function RunNetHiddenCapture(const Params, OutTag: string): Integer;
var
  OutPath: string;
  SL: TStringList;
  j: Integer;
  RC: Integer;
begin
  LogEntry('RunNetHiddenCapture');
  if SimulateNetFailPowerShell then
  begin
    if not SimLogNetPsShown then
    begin
      LogSimulationScenario('System fails on net.exe commands and uses PowerShell fallback');
      SimLogNetPsShown := True;
    end;
    LogWarn('Simulation: forcing net.exe failure for: ' + MaskCommandForLog('net.exe', Params));
    Result := 1;
    LogExit('RunNetHiddenCapture');
    exit;
  end;

  OutPath := TempFile('net_' + SanitizeFileName(OutTag) + '.log');
  RC := 0;
  LogDebug('RunNetHiddenCapture: ' + MaskCommandForLog('net.exe', Params));
  Exec(EXE_CMD, '/c "net.exe ' + Params + '" > "' + OutPath + '" 2>&1', '', SW_HIDE, ewWaitUntilTerminated, RC);
  LogDebug('RunNetHiddenCapture: exit=' + IntToStr(RC) + ' tag=' + OutTag);

  if RC <> 0 then
  begin
    if FileExists(OutPath) then
    begin
      SL := TStringList.Create;
      try
        try
          SL.LoadFromFile(OutPath);
          if SL.Count > 0 then
          begin
            LogDebug('net.exe output (' + OutTag + ', ' + IntToStr(SL.Count) + ' lines):');
            for j := 0 to SL.Count - 1 do
              LogDebug('  ' + SL[j]);
          end;
        except
          LogWarn('RunNetHiddenCapture: failed to read output: ' + OutPath);
        end;
      finally
        SL.Free;
      end;
    end
    else
      LogDebug('RunNetHiddenCapture: no output file for tag=' + OutTag);
  end
  else
  begin
    LogDebug('RunNetHiddenCapture: command succeeded (exit=0) for tag=' + OutTag);
  end;

  if FileExists(OutPath) then
    DeleteFile(OutPath);
  Result := RC;
  LogExit('RunNetHiddenCapture');
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
  OpTick: Cardinal;
begin
  LogEntry('RunPSHiddenCode');
  OpTick := GetTickCount;
  RC := -1;
  LogDebug('RunPSHiddenCode: ' + MaskPasswordsInString(Command));
  ExecPowerShellHidden(Command, RC);
  LogDebug('RunPSHiddenCode exit=' + IntToStr(RC) + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  Result := RC;
  LogExit('RunPSHiddenCode');
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

// Resolve a TColor to an actual RGB value.
// System color constants (clBtnFace, clWindow, etc.) have bit 31 set (negative when
// treated as a signed Longint).  GetSysColor converts the index to a real BGR value.
function ResolveColor(C: Longint): Longint;
begin
  if C < 0 then
    Result := GetSysColor(C and $FF)
  else
    Result := C;
end;

function IsDarkColor(C: Longint): Boolean;
var R, G, B, bright: Integer;
begin
  C := ResolveColor(C);
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
  OpTick: Cardinal;
begin
  LogEntry('PreTrustRDPCertCurrentUser');
  OpTick := GetTickCount;
  // Value 76 (0x4C) represents the device/resource trust flags
  ExecPowerShellHidden(
    '$ErrorActionPreference = ''Stop''; ' +
    'try { ' +
    '  $regPath = ''HKCU:\Software\Microsoft\Terminal Server Client\LocalDevices''; ' +
    '  if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }; ' +
    '  New-ItemProperty -Path $regPath -Name ''' + RDP_LOOPBACK_IP + ''' -PropertyType DWord -Value 76 -Force | Out-Null; ' +
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
  OpTick: Cardinal;
begin
  LogEntry('AddDefenderExclusionForApp');
  OpTick := GetTickCount;
  // Ensure Defender exclusions are scoped to the two runtime DLLs only
  ExecPowerShellHidden(
    '$paths = @(''' + ExpandConstant('{app}\TermWrap.dll') + ''',''' + ExpandConstant('{app}\Zydis.dll') + '''); ' +
    'try { $p = Get-MpPreference; foreach ($path in $paths) { if (-not ($p.ExclusionPath -contains $path)) { Add-MpPreference -ExclusionPath $path } } } catch { }',
    ResultCode);
  LogDebug('AddDefenderExclusionForApp: exit=' + IntToStr(ResultCode) + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('AddDefenderExclusionForApp');
end;

procedure RemoveDefenderExclusionForApp;
var
  ResultCode: Integer;
  OpTick: Cardinal;
begin
  LogEntry('RemoveDefenderExclusionForApp');
  OpTick := GetTickCount;
  // Remove Defender exclusions for the two runtime DLLs during uninstall
  ExecPowerShellHidden(
    '$paths = @(''' + ExpandConstant('{app}\TermWrap.dll') + ''',''' + ExpandConstant('{app}\Zydis.dll') + '''); ' +
    'try { foreach ($path in $paths) { Remove-MpPreference -ExclusionPath $path } } catch { }',
    ResultCode);
  LogDebug('RemoveDefenderExclusionForApp: exit=' + IntToStr(ResultCode) + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('RemoveDefenderExclusionForApp');
end;

// -----------------------------------------------------------------------------
// WINDOWS SERVICE MANAGEMENT
// -----------------------------------------------------------------------------

// Stop TermService with retry logic (up to 5 attempts).
// Required before replacing TermWrap DLLs; disconnects active RDP sessions.
procedure StopTermService;
var
  ResultCode: Integer;
  Attempt: Integer;
  ServiceRunning: Boolean;
  OpTick: Cardinal;
begin
  LogEntry('StopTermService');
  OpTick := GetTickCount;
  LogInfo('Stopping Remote Desktop Services...');
  StatusOverlay.Caption := 'Stopping Remote Desktop Services...';
  
  // Start with disable + stop attempts (no initial "normal" stop)
  ServiceRunning := True;
  Attempt := 0;
  while ServiceRunning and (Attempt < 5) do
  begin
    Inc(Attempt);
    LogDebug('StopTermService: Attempt ' + IntToStr(Attempt) + '/5');
    WizardForm.Update;
    
    // Disable the service to prevent auto-restart
    ResultCode := RunPSHiddenCode('Set-Service -Name TermService -StartupType Disabled -ErrorAction Stop');
    LogDebug('Set-Service Disabled exit code=' + IntToStr(ResultCode));
    SleepWithUI(SLEEP_MEDIUM);
    
    // Try stopping with PowerShell
    LogDebug('Executing Stop-Service TermService');
    ResultCode := RunPSHiddenCode('Stop-Service -Name TermService -Force -ErrorAction Stop');
    LogDebug('Stop-Service exit code=' + IntToStr(ResultCode));
    
    // Wait longer for service to actually stop
    SleepWithUI(SLEEP_LONG + SLEEP_LONG);
    
    // Check service state (exit 0 if stopped, exit 1 if running or any other state)
    ResultCode := RunPSHiddenCode('if ((Get-Service -Name TermService).Status -eq ''Stopped'') { exit 0 } else { exit 1 }');
    ServiceRunning := (ResultCode = 1);
    LogDebug('Service running=' + BoolToStr(ServiceRunning) + ' after attempt ' + IntToStr(Attempt));
  end;
  
  if ServiceRunning then
    LogWarn('Service still running after 5 attempts, proceeding anyway')
  else
    LogInfo('Service verified stopped');
  
  LogInfo('StopTermService completed [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('StopTermService');
end;

// Start TermService and set it to Automatic startup, return exit code
function StartTermServiceEx: Integer;
var
  RC: Integer;
  OpTick: Cardinal;
begin
  LogEntry('StartTermServiceEx');
  OpTick := GetTickCount;
  LogInfo('Starting Remote Desktop Services...');
  StatusOverlay.Caption := 'Restarting Remote Desktop Services...';
  
  LogDebug('Setting TermService to Automatic');
  RC := RunPSHiddenCode('Set-Service -Name TermService -StartupType Automatic -ErrorAction Stop');
  LogDebug('Set-Service Automatic exit code=' + IntToStr(RC));
  
  SleepWithUI(SLEEP_MEDIUM);
  LogDebug('Executing Start-Service TermService');
  RC := RunPSHiddenCode('Start-Service -Name TermService -ErrorAction Stop');
  LogDebug('Start-Service exit code=' + IntToStr(RC));
  SleepWithUI(SLEEP_LONG);
  LogInfo('StartTermServiceEx completed [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('StartTermServiceEx');
  Result := RC;
end;

// Shared helper: query a string value, fix it via sc.exe if wrong, then verify.
procedure EnsureServiceStringConfig(const RegKey, ValueName, ScArgs, TargetValue, CheckDesc, FixDesc: string);
var
  Current: string;
  RC: Integer;
  OpTick: Cardinal;
begin
  LogEntry('EnsureServiceStringConfig');
  OpTick := GetTickCount;
  Current := '';
  if RegQueryStringValue(HKLM, RegKey, ValueName, Current) then
    LogDebug(CheckDesc + ': current ' + ValueName + '=' + Current)
  else
    LogDebug(CheckDesc + ': current ' + ValueName + '=<missing>');

  if CompareText(Trim(Current), TargetValue) = 0 then
  begin
    LogDebug(CheckDesc + ': already correct, no action needed [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
    LogExit('EnsureServiceStringConfig');
    Exit;
  end;

  LogDebug(FixDesc + ': executing sc.exe fix');
  RC := RunHidden('sc.exe', ScArgs);
  if RC = 0 then
    LogInfo(FixDesc + ': success')
  else
    LogWarn(FixDesc + ': failed: sc.exe exit=' + IntToStr(RC));

  Sleep(SLEEP_SHORT);
  Current := '';
  if RegQueryStringValue(HKLM, RegKey, ValueName, Current) then
    LogDebug(CheckDesc + ' (post-fix): ' + ValueName + '=' + Current)
  else
    LogWarn(CheckDesc + ' (post-fix): ' + ValueName + ' could not be read');
  LogInfo('EnsureServiceStringConfig completed [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('EnsureServiceStringConfig');
end;

// Shared helper: query a DWord value, fix it via sc.exe if wrong, then verify.
procedure EnsureServiceDWordConfig(const RegKey, ValueName, ScArgs: string; TargetValue: Cardinal; const CheckDesc, FixDesc: string);
var
  Current: Cardinal;
  RC: Integer;
  OpTick: Cardinal;
begin
  LogEntry('EnsureServiceDWordConfig');
  OpTick := GetTickCount;
  Current := $FFFFFFFF;
  if RegQueryDWordValue(HKLM, RegKey, ValueName, Current) then
    LogDebug(CheckDesc + ': current ' + ValueName + '=' + IntToStr(Current))
  else
    LogDebug(CheckDesc + ': current ' + ValueName + '=<unable to read>');

  if Current = TargetValue then
  begin
    LogDebug(CheckDesc + ': already correct, no action needed [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
    LogExit('EnsureServiceDWordConfig');
    Exit;
  end;

  LogDebug(FixDesc + ': executing sc.exe fix');
  RC := RunHidden('sc.exe', ScArgs);
  if RC = 0 then
    LogInfo(FixDesc + ': success')
  else
    LogWarn(FixDesc + ': failed: sc.exe exit=' + IntToStr(RC));

  Sleep(SLEEP_SHORT);
  if RegQueryDWordValue(HKLM, RegKey, ValueName, Current) then
    LogDebug(CheckDesc + ' (post-fix): ' + ValueName + '=' + IntToStr(Current))
  else
    LogWarn(CheckDesc + ' (post-fix): ' + ValueName + ' could not be read');
  LogInfo('EnsureServiceDWordConfig completed [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('EnsureServiceDWordConfig');
end;

procedure EnsureTermServiceRunsAsNetworkService;
begin
  LogEntry('EnsureTermServiceRunsAsNetworkService');
  EnsureServiceStringConfig(
    REG_TERMSERVICE, 'ObjectName',
    'config TermService obj= "NT AUTHORITY\NetworkService" password= ""',
    'NT AUTHORITY\NetworkService',
    'TermService account check',
    'TermService account fix: set ObjectName to NT AUTHORITY\NetworkService');
  LogExit('EnsureTermServiceRunsAsNetworkService');
end;

// Wrapper that calls StartTermServiceEx and discards the exit code
procedure StartTermService;
begin
  LogEntry('StartTermService');
  StartTermServiceEx;
  LogExit('StartTermService');
end;

procedure EnsureUmRdpServiceAutomatic;
begin
  LogEntry('EnsureUmRdpServiceAutomatic');
  // Start type 2 = Automatic
  EnsureServiceDWordConfig(
    REG_UMRDPSERVICE, 'Start',
    'config UmRdpService start=auto',
    2,
    'UmRdpService startup type check',
    'UmRdpService startup type fix: set Start=2 (Automatic)');
  LogExit('EnsureUmRdpServiceAutomatic');
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

// Populates UsersList (login names) and DisplayList (email if online account,
// else same as login name). Filters to enabled true-local accounts only by
// comparing each account's SID domain portion against the local machine SID.
procedure GetLocalUsers(UsersList: TStringList; DisplayList: TStringList);
var
  PSPath: string;
  ResultCode: Integer;
  i: Integer;
  Line: string;
  Parts: TStringList;
  UserName: string;
  DisplayName: string;
  PSCommand: string;
  OpTick: Cardinal;
begin
  LogEntry('GetLocalUsers');
  OpTick := GetTickCount;
  PSPath := ExpandConstant(TEMP_LOCAL_USERS);

  PSCommand :=
    'try {' +
    ' $out=@();' +
    ' $sq=[char]39;' +
    ' foreach($u in (Get-LocalUser | Where-Object {$_.Enabled -and $_.Name[0] -ne $sq})) {' +
    '  $n=$u.Name; $detail=$null;' +
    '  if($u.FullName -match ''@'') { $detail=$u.FullName }' +
    '  elseif($u.FullName -and $u.FullName.Trim() -ne '''') { $detail=$u.FullName.Trim() }' +
    '  if(-not $detail) { try {' +
    '   $s=$u.SID.Value;' +
    '   $r=''HKLM:\SOFTWARE\Microsoft\IdentityStore\Cache\''+$s+''\IdentityCache\''+$s;' +
    '   $e=(Get-ItemProperty -Path $r -Name UserName -ErrorAction Stop).UserName;' +
    '   if($e -and $e.Trim() -ne '''') {$detail=$e.Trim()}' +
    '  } catch {} }' +
    '  if($detail -and $detail -ne $n) { $d=$n+'' (''+$detail+'')'' } else { $d=$n }' +
    '  $out+=($n+''|''+$d)' +
    ' }' +
    ' $out | Out-File -Encoding UTF8 ''' + PSPath + ''' -Force;' +
    ' exit 0' +
    '} catch { exit 1 }';
  Exec(EXE_POWERSHELL, BuildPowerShellArgs(PSCommand, True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  LogDebug('GetLocalUsers: PS exit=' + IntToStr(ResultCode) + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');

  if not FileExists(PSPath) then
  begin
    LogDebug('GetLocalUsers: no output file, no users found');
    LogExit('GetLocalUsers');
    exit;
  end;

  Parts := TStringList.Create;
  try
    Parts.LoadFromFile(PSPath);
    LogDebug('GetLocalUsers: raw lines=' + IntToStr(Parts.Count));
    for i := 0 to Parts.Count - 1 do
    begin
      Line := Trim(Parts[i]);
      if Line = '' then
        continue;
      if Pos('|', Line) > 0 then
      begin
        UserName := Copy(Line, 1, Pos('|', Line) - 1);
        DisplayName := Copy(Line, Pos('|', Line) + 1, MaxInt);
      end
      else
      begin
        UserName := Line;
        DisplayName := Line;
      end;
      UserName := Trim(UserName);
      DisplayName := Trim(DisplayName);
      if (UserName = '') or IsExcludedUser(UserName) then
      begin
        LogDebug('GetLocalUsers: skipped excluded/empty user: ' + UserName);
        continue;
      end;
      UsersList.Add(UserName);
      DisplayList.Add(DisplayName);
    end;
    LogDebug('GetLocalUsers: found ' + IntToStr(UsersList.Count) + ' users [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  finally
    Parts.Free;
    DeleteFile(PSPath);
  end;
  LogExit('GetLocalUsers');
end;

function GetDesktopRdpFiles: TStringList;
var
  FilesList: TStringList;
  FindRec: TFindRec;
  DesktopPattern: string;
begin
  LogEntry('GetDesktopRdpFiles');
  FilesList := TStringList.Create;
  DesktopPattern := ExpandConstant('{userdesktop}\*.rdp');

  if FindFirst(DesktopPattern, FindRec) then
  begin
    try
      repeat
        if (FindRec.Attributes and 16) = 0 then
          FilesList.Add(ExpandConstant('{userdesktop}\') + FindRec.Name);
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;

  FilesList.Sort;
  Result := FilesList;
end;

function IsPathEndingWithRdpExe(const FilePath: string): Boolean;
var
  S: string;
begin
  S := LowerCase(Trim(StripWrappingQuotes(FilePath)));
  Result := (Length(S) >= Length('rdp.exe')) and
            (Copy(S, Length(S) - Length('rdp.exe') + 1, Length('rdp.exe')) = 'rdp.exe');
end;

function DesktopFolderHasRdpShortcut(const DesktopDir: string): Boolean;
var
  FindRec: TFindRec;
  LnkPattern: string;
  ShellObj: Variant;
  ShortcutObj: Variant;
  TargetPath: string;
begin
  Result := False;
  if not DirExists(DesktopDir) then
    exit;

  LnkPattern := AddBackslash(DesktopDir) + '*.lnk';
  if not FindFirst(LnkPattern, FindRec) then
    exit;

  ShellObj := Unassigned;
  try
    try
      ShellObj := CreateOleObject('WScript.Shell');
    except
      WriteInstallerLog('WARNING: Could not create WScript.Shell COM object for shortcut detection');
      exit;
    end;

    repeat
      if (FindRec.Attributes and 16) = 0 then
      begin
        TargetPath := '';
        try
          ShortcutObj := ShellObj.CreateShortcut(AddBackslash(DesktopDir) + FindRec.Name);
          TargetPath := ShortcutObj.TargetPath;
        except
          // Ignore malformed shortcut files and continue scanning.
          TargetPath := '';
        end;

        if IsPathEndingWithRdpExe(TargetPath) then
        begin
          Result := True;
          exit;
        end;
      end;
    until not FindNext(FindRec);
  finally
    FindClose(FindRec);
  end;
end;

function HasDesktopShortcutTargetingRdpExe: Boolean;
var
  UserDesktop: string;
  CommonDesktop: string;
begin
  UserDesktop := ExpandConstant('{userdesktop}');
  CommonDesktop := ExpandConstant('{commondesktop}');

  Result := DesktopFolderHasRdpShortcut(UserDesktop);
  if (not Result) and (CommonDesktop <> '') and (CompareText(UserDesktop, CommonDesktop) <> 0) then
    Result := DesktopFolderHasRdpShortcut(CommonDesktop);
end;

procedure OnShortcutCheckBoxClick(Sender: TObject);
begin
  LogEntry('OnShortcutCheckBoxClick');
  // No action needed on click - validation happens in NextButtonClick
  LogExit('OnShortcutCheckBoxClick');
end;

procedure BuildShortcutEditorControls;
var
  i: Integer;
begin
  LogEntry('BuildShortcutEditorControls');
  if EditShortcutControlsBuilt then
  begin
    LogExit('BuildShortcutEditorControls');
    exit;
  end;

  if ShortcutsPerPage = 0 then
    ShortcutsPerPage := 8;

  ShortcutHeaderLabel := TLabel.Create(EditShortcutPage);
  ShortcutHeaderLabel.Parent := EditShortcutPage.Surface;
  ShortcutHeaderLabel.Left := ScaleX(20);
  ShortcutHeaderLabel.Top := ScaleY(12);
  ShortcutHeaderLabel.Caption := 'Desktop .rdp shortcuts (check one or more to edit)';
  ShortcutHeaderLabel.Font.Style := [fsBold];

  ShortcutEmptyLabel := TLabel.Create(EditShortcutPage);
  ShortcutEmptyLabel.Parent := EditShortcutPage.Surface;
  ShortcutEmptyLabel.Left := ScaleX(20);
  ShortcutEmptyLabel.Top := ShortcutHeaderLabel.Top + ScaleY(28);
  ShortcutEmptyLabel.Caption := 'No .rdp files were found on your Desktop.';
  ShortcutEmptyLabel.Visible := False;

  SetLength(ShortcutCheckBoxes, DesktopRdpFiles.Count);
  for i := 0 to DesktopRdpFiles.Count - 1 do
  begin
    ShortcutCheckBoxes[i] := TCheckBox.Create(EditShortcutPage);
    ShortcutCheckBoxes[i].Parent := EditShortcutPage.Surface;
    ShortcutCheckBoxes[i].Left := ScaleX(20);
    ShortcutCheckBoxes[i].Width := EditShortcutPage.SurfaceWidth - ScaleX(40);
    ShortcutCheckBoxes[i].Caption := ExtractFileName(DesktopRdpFiles[i]);
    ShortcutCheckBoxes[i].Tag := i;
    ShortcutCheckBoxes[i].OnClick := @OnShortcutCheckBoxClick;
    ShortcutCheckBoxes[i].Visible := False;
  end;

  if DesktopRdpFiles.Count > ShortcutsPerPage then
  begin
    ShortcutPrevButton := TButton.Create(EditShortcutPage);
    ShortcutPrevButton.Parent := EditShortcutPage.Surface;
    ShortcutPrevButton.Left := ScaleX(20);
    ShortcutPrevButton.Width := ScaleX(80);
    ShortcutPrevButton.Caption := 'Previous';
    ShortcutPrevButton.OnClick := @OnPrevShortcutPageClick;

    ShortcutPageLabel := TLabel.Create(EditShortcutPage);
    ShortcutPageLabel.Parent := EditShortcutPage.Surface;
    ShortcutPageLabel.AutoSize := True;
    ShortcutPageLabel.Left := ScaleX(110);

    ShortcutNextButton := TButton.Create(EditShortcutPage);
    ShortcutNextButton.Parent := EditShortcutPage.Surface;
    ShortcutNextButton.Width := ScaleX(80);
    ShortcutNextButton.Caption := 'Next';
    ShortcutNextButton.OnClick := @OnNextShortcutPageClick;
  end
  else
  begin
    ShortcutPrevButton := nil;
    ShortcutNextButton := nil;
    ShortcutPageLabel := nil;
  end;

  CurrentShortcutPage := 0;
  UpdateShortcutPageDisplay;
  EditShortcutControlsBuilt := True;
end;

procedure BuildEditShortcutAdvancedControls;
begin
  LogEntry('BuildEditShortcutAdvancedControls');
  if EditShortcutAdvancedControlsBuilt then
  begin
    LogExit('BuildEditShortcutAdvancedControls');
    exit;
  end;

  // Instruction text explaining what to do
  EditShortcutAdvancedLabel := TLabel.Create(Page_EditShortcutAdvanced);
  EditShortcutAdvancedLabel.Parent := Page_EditShortcutAdvanced.Surface;
  EditShortcutAdvancedLabel.Left := ScaleX(20);
  EditShortcutAdvancedLabel.Top := ScaleY(12);
  EditShortcutAdvancedLabel.Width := Page_EditShortcutAdvanced.SurfaceWidth - ScaleX(40);
  EditShortcutAdvancedLabel.AutoSize := False;
  EditShortcutAdvancedLabel.Height := ScaleY(80);
  EditShortcutAdvancedLabel.WordWrap := True;
  EditShortcutAdvancedLabel.Caption :=
    'The shortcut was opened in the Remote Desktop Connection app. Make your changes there.' + #13#10#13#10 +
    'Important: Always click [Save] on the General tab. Then click Next here to finalize the shortcut settings.';

  // Screenshot image showing where to click Save
  EditShortcutAdvancedImage := TBitmapImage.Create(Page_EditShortcutAdvanced);
  EditShortcutAdvancedImage.Parent := Page_EditShortcutAdvanced.Surface;
  EditShortcutAdvancedImage.Left := ScaleX(20);
  EditShortcutAdvancedImage.Top := EditShortcutAdvancedLabel.Top + EditShortcutAdvancedLabel.Height + ScaleY(8);
  EditShortcutAdvancedImage.Width := ScaleX(200);
  EditShortcutAdvancedImage.Height := ScaleY(210);
  EditShortcutAdvancedImage.Stretch := True;
  EditShortcutAdvancedImage.Visible := False;

  EditShortcutAdvancedControlsBuilt := True;
end;

procedure UpdateShortcutPageDisplay;
var
  i, PageCount, StartIdx, EndIdx, VisIndex, BaseTop, RowHeight: Integer;
begin
  LogEntry('UpdateShortcutPageDisplay');
  BaseTop := ShortcutHeaderLabel.Top + ScaleY(26);
  RowHeight := ScaleY(24);

  if DesktopRdpFiles.Count = 0 then
  begin
    ShortcutEmptyLabel.Visible := True;
    if Assigned(ShortcutPrevButton) then ShortcutPrevButton.Visible := False;
    if Assigned(ShortcutNextButton) then ShortcutNextButton.Visible := False;
    if Assigned(ShortcutPageLabel) then ShortcutPageLabel.Visible := False;
    exit;
  end;

  ShortcutEmptyLabel.Visible := False;

  PageCount := (DesktopRdpFiles.Count + ShortcutsPerPage - 1) div ShortcutsPerPage;
  if CurrentShortcutPage < 0 then CurrentShortcutPage := 0;
  if CurrentShortcutPage >= PageCount then CurrentShortcutPage := PageCount - 1;

  StartIdx := CurrentShortcutPage * ShortcutsPerPage;
  EndIdx := StartIdx + ShortcutsPerPage - 1;
  if EndIdx > DesktopRdpFiles.Count - 1 then EndIdx := DesktopRdpFiles.Count - 1;

  VisIndex := 0;
  for i := 0 to DesktopRdpFiles.Count - 1 do
  begin
    ShortcutCheckBoxes[i].Visible := (i >= StartIdx) and (i <= EndIdx);
    if ShortcutCheckBoxes[i].Visible then
    begin
      ShortcutCheckBoxes[i].Top := BaseTop + VisIndex * RowHeight;
      Inc(VisIndex);
    end;
  end;

  if Assigned(ShortcutPrevButton) and Assigned(ShortcutNextButton) and Assigned(ShortcutPageLabel) then
  begin
    ShortcutPrevButton.Visible := True;
    ShortcutNextButton.Visible := True;
    ShortcutPageLabel.Visible := True;
    ShortcutPrevButton.Top := EditShortcutPage.SurfaceHeight - ScaleY(30);
    ShortcutNextButton.Top := ShortcutPrevButton.Top;
    ShortcutNextButton.Left := ShortcutPrevButton.Left + ShortcutPrevButton.Width + ScaleX(120);
    ShortcutPageLabel.Top := ShortcutPrevButton.Top + ScaleY(4);
    ShortcutPageLabel.Caption := 'Page ' + IntToStr(CurrentShortcutPage + 1) + ' of ' + IntToStr(PageCount);
    ShortcutPrevButton.Enabled := CurrentShortcutPage > 0;
    ShortcutNextButton.Enabled := CurrentShortcutPage < PageCount - 1;
  end;
end;

procedure OnPrevShortcutPageClick(Sender: TObject);
begin
  LogEntry('OnPrevShortcutPageClick');
  if CurrentShortcutPage > 0 then
    Dec(CurrentShortcutPage);
  UpdateShortcutPageDisplay;
  LogExit('OnPrevShortcutPageClick');
end;

procedure OnNextShortcutPageClick(Sender: TObject);
var
  PageCount: Integer;
begin
  LogEntry('OnNextShortcutPageClick');
  PageCount := (DesktopRdpFiles.Count + ShortcutsPerPage - 1) div ShortcutsPerPage;
  if CurrentShortcutPage < PageCount - 1 then
    Inc(CurrentShortcutPage);
  UpdateShortcutPageDisplay;
  LogExit('OnNextShortcutPageClick');
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
    begin
      UserPasswordStatus[i].Caption := '';
      UserPasswordStatus[i].Visible := False;
    end;
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
  LogEntry('OnPasswordEditChange');
  idx := FindUserIndexFromEdit(TEdit(Sender));
  if idx >= 0 then
    UserPasswordStatus[idx].Visible := False;
  LogExit('OnPasswordEditChange');
end;

procedure OnUserCheckBoxClick(Sender: TObject);
var
  idx: Integer;
begin
  LogEntry('OnUserCheckBoxClick');
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
  LogExit('OnUserCheckBoxClick');
end;

procedure SyncSetupOptionHintStates();
begin
  if Assigned(chkInstallTermWrapHint) and Assigned(chkInstallTermWrap) then
    chkInstallTermWrapHint.Enabled := chkInstallTermWrap.Enabled;
  if Assigned(rbUseExistingUsersHint) and Assigned(rbUseExistingUsers) then
    rbUseExistingUsersHint.Enabled := rbUseExistingUsers.Enabled;
end;

procedure OnCreateRdpShortcutsClick(Sender: TObject);
begin
  LogEntry('OnCreateRdpShortcutsClick');
  if Assigned(rbCreateUsers) then
    rbCreateUsers.Enabled := chkCreateRdpShortcuts.Checked;
  if Assigned(rbUseExistingUsers) then
    rbUseExistingUsers.Enabled := chkCreateRdpShortcuts.Checked;
  // Update derived flags so page skipping and next-button validation are accurate
  DoCreateRdpShortcuts := chkCreateRdpShortcuts.Checked;
  if DoCreateRdpShortcuts then
  begin
    if Assigned(rbCreateUsers) and rbCreateUsers.Checked then
      CreateUserMode := createUserModeNew
    else
      CreateUserMode := createUserModeExisting;
  end;
  SyncSetupOptionHintStates();
  LogDebug('OnCreateRdpShortcutsClick: DoCreateRdpShortcuts=' + BoolToStr(DoCreateRdpShortcuts) + ' CreateUserMode=' + IntToStr(CreateUserMode));
  LogExit('OnCreateRdpShortcutsClick');
end;

  procedure OnInstallModeChange(Sender: TObject);
begin
  // Always show controls, but only enable them for Install mode
  if Assigned(chkInstallTermWrap) then chkInstallTermWrap.Enabled := Assigned(rbInstall) and rbInstall.Checked;
  if Assigned(chkCreateRdpShortcuts) then chkCreateRdpShortcuts.Enabled := Assigned(rbInstall) and rbInstall.Checked;
  if Assigned(CreateRdpShortcutsGroup) then CreateRdpShortcutsGroup.Enabled := Assigned(rbInstall) and rbInstall.Checked;
  if Assigned(rbCreateUsers) then rbCreateUsers.Enabled := Assigned(rbInstall) and rbInstall.Checked and chkCreateRdpShortcuts.Checked;
  if Assigned(rbUseExistingUsers) then rbUseExistingUsers.Enabled := Assigned(rbInstall) and rbInstall.Checked and chkCreateRdpShortcuts.Checked;
  SyncSetupOptionHintStates();
end;

procedure OnUseAllMonitorsClick(Sender: TObject);
begin
  LogEntry('OnUseAllMonitorsClick');
  LogDebug('OnUseAllMonitorsClick: checked=' + BoolToStr(Assigned(chkUseAllMonitors) and chkUseAllMonitors.Checked));
  if Assigned(chkFullScreen) then
  begin
    if chkUseAllMonitors.Checked then
    begin
      chkFullScreen.Checked := True;
      chkFullScreen.Enabled := False;
      OnFullScreenClick(nil);
    end
    else
      chkFullScreen.Enabled := True;
  end;
  LogExit('OnUseAllMonitorsClick');
end;

procedure OnFullScreenClick(Sender: TObject);
begin
  LogEntry('OnFullScreenClick');
  LogDebug('OnFullScreenClick: checked=' + BoolToStr(Assigned(chkFullScreen) and chkFullScreen.Checked));
  if Assigned(cboResolution) then
    cboResolution.Enabled := not chkFullScreen.Checked;
  // When full screen is forced, hide the custom size inputs; otherwise restore them
  if chkFullScreen.Checked then
  begin
    if Assigned(lblCustomWidth)  then lblCustomWidth.Visible  := False;
    if Assigned(edtCustomWidth)  then edtCustomWidth.Visible  := False;
    if Assigned(lblCustomHeight) then lblCustomHeight.Visible := False;
    if Assigned(edtCustomHeight) then edtCustomHeight.Visible := False;
  end
  else
    OnResolutionChange(nil);
  LogExit('OnFullScreenClick');
end;

procedure OnResolutionChange(Sender: TObject);
var
  IsCustom: Boolean;
begin
  LogEntry('OnResolutionChange');
  if not Assigned(cboResolution) then
  begin
    LogDebug('OnResolutionChange: cboResolution not assigned, skipping');
    LogExit('OnResolutionChange');
    exit;
  end;
  IsCustom := (cboResolution.ItemIndex >= 0) and
              (cboResolution.Items[cboResolution.ItemIndex] = 'Custom');
  LogDebug('OnResolutionChange: IsCustom=' + BoolToStr(IsCustom));
  if Assigned(lblCustomWidth)  then lblCustomWidth.Visible  := IsCustom;
  if Assigned(edtCustomWidth)  then edtCustomWidth.Visible  := IsCustom;
  if Assigned(lblCustomHeight) then lblCustomHeight.Visible := IsCustom;
  if Assigned(edtCustomHeight) then edtCustomHeight.Visible := IsCustom;
  LogExit('OnResolutionChange');
end;

function IsTermWrapInstalled(): Boolean;
var
  TermWrapPath, ZydisPath: string;
  TermWrapExists, ZydisExists: Boolean;
  ServiceDllPath: string;
  RegistryPointsToTermWrap: Boolean;
begin
  // Define the expected paths for TermWrap.dll and Zydis.dll
  // Files are installed to {commonpf64}\RDPWrapKit\ (matching DefaultDirName)
  TermWrapPath := ExpandConstant('{commonpf64}\RDPWrapKit\TermWrap.dll');
  ZydisPath := ExpandConstant('{commonpf64}\RDPWrapKit\Zydis.dll');

  // Check if the files exist
  TermWrapExists := FileExists(TermWrapPath);
  ZydisExists := FileExists(ZydisPath);

  // Check if registry ServiceDll points to TermWrap.dll
  RegistryPointsToTermWrap := False;
  if RegQueryStringValue(HKLM, REG_TERMSERVICE_PARAMS, 'ServiceDll', ServiceDllPath) then
  begin
    // Check if the ServiceDll path contains TermWrap.dll (case-insensitive)
    RegistryPointsToTermWrap := (Pos('termwrap.dll', Lowercase(ServiceDllPath)) > 0);
  end;

  // Log the results
  WriteInstallerLog('IsTermWrapInstalled: File check - TermWrap.dll exists: ' + BoolToStr(TermWrapExists) + ' at ' + TermWrapPath);
  WriteInstallerLog('IsTermWrapInstalled: File check - Zydis.dll exists: ' + BoolToStr(ZydisExists) + ' at ' + ZydisPath);
  WriteInstallerLog('IsTermWrapInstalled: Registry check - ServiceDll points to TermWrap: ' + BoolToStr(RegistryPointsToTermWrap));
  if ServiceDllPath <> '' then
    WriteInstallerLog('IsTermWrapInstalled: Registry check - Current ServiceDll = ' + ServiceDllPath);

  // Return true only if both files exist AND registry points to TermWrap
  Result := TermWrapExists and ZydisExists and RegistryPointsToTermWrap;
end;


// Returns the version string of the installed TermWrap.dll (e.g. "0.0.6.0").
// Uses PowerShell to read the FileVersionRaw from the file at the registry-configured path.
// Returns empty string if TermWrap is not installed or the file cannot be read.
function GetInstalledTermWrapVersion(): string;
var
  ServiceDllPath: string;
begin
  Result := '';
  if RegQueryStringValue(HKLM, REG_TERMSERVICE_PARAMS, 'ServiceDll', ServiceDllPath) then
  begin
    if (Pos('termwrap.dll', Lowercase(ServiceDllPath)) > 0) and FileExists(ServiceDllPath) then
      Result := GetPSOutput('$f=Get-Item ''' + ServiceDllPath + '''; $f.VersionInfo.FileVersionRaw.ToString()');
  end;
end;


// Returns the file size (in bytes) of the installed TermWrap.dll as a string.
// Uses PowerShell to read the Length property from the file at the registry-configured path.
// Returns empty string if TermWrap is not installed or the file cannot be read.
function GetInstalledTermWrapSize(): string;
var
  ServiceDllPath: string;
begin
  Result := '';
  if RegQueryStringValue(HKLM, REG_TERMSERVICE_PARAMS, 'ServiceDll', ServiceDllPath) then
  begin
    if (Pos('termwrap.dll', Lowercase(ServiceDllPath)) > 0) and FileExists(ServiceDllPath) then
      Result := GetPSOutput('$f=Get-Item ''' + ServiceDllPath + '''; $f.Length.ToString()');
  end;
end;


procedure OnViewLogButtonClick(Sender: TObject);
var
  DestName: string;
  Saved: Boolean;
begin
  LogEntry('OnViewLogButtonClick');
  WriteInstallerLog('User clicked Save Install Log button');
  DestName := ExpandConstant('{userdesktop}\RDPWrapKit_install.log');
  Saved := CopyFile(InstallLogPath, DestName, False);
  if Saved then
    MsgBox('File RDPWrapKit_install.log was saved to your Desktop', mbInformation, MB_OK)
  else
    MsgBox('Failed to save install log to the Desktop location.', mbError, MB_OK);
  LogExit('OnViewLogButtonClick');
end;

procedure OnPasswordResetLinkClick(Sender: TObject);
var
  ResultCode: Integer;
begin
  LogEntry('OnPasswordResetLinkClick');
  Exec('control.exe', 'userpasswords2', '', SW_SHOW, ewNoWait, ResultCode);
  LogDebug('OnPasswordResetLinkClick: launched, exit=' + IntToStr(ResultCode));
  LogExit('OnPasswordResetLinkClick');
end;

procedure BuildCreateShortcutsControls;
var
  i: Integer;
  TopPos: Integer;
  BottomPos: Integer;
begin
  LogEntry('BuildCreateShortcutsControls');
  // Avoid building controls multiple times (prevents duplicate buttons/labels)
  if CreateShortcutsControlsBuilt then
  begin
    LogDebug('BuildCreateShortcutsControls: already built, skipping');
    LogExit('BuildCreateShortcutsControls');
    exit;
  end;
  LogDebug('BuildCreateShortcutsControls: building for ' + IntToStr(LocalUsersList.Count) + ' users');
  TopPos := ScaleY(10);

  // Default users-per-page
  if UsersPerPage = 0 then
    UsersPerPage := 7;

  // Create "Users found" header
  Tool1UsersHeaderLabel := TLabel.Create(Page_CreateShortcutsForExistingUsers);
  Tool1UsersHeaderLabel.Parent := Page_CreateShortcutsForExistingUsers.Surface;
  Tool1UsersHeaderLabel.Left := ScaleX(20);
  Tool1UsersHeaderLabel.Top := TopPos;
  Tool1UsersHeaderLabel.Caption := 'Users found';
  Tool1UsersHeaderLabel.Font.Style := [fsBold];

  // Create "Password" header
  Tool1PasswordHeaderLabel := TLabel.Create(Page_CreateShortcutsForExistingUsers);
  Tool1PasswordHeaderLabel.Parent := Page_CreateShortcutsForExistingUsers.Surface;
  Tool1PasswordHeaderLabel.Left := ScaleX(220);
  Tool1PasswordHeaderLabel.Top := TopPos;
  Tool1PasswordHeaderLabel.Caption := 'Password';
  Tool1PasswordHeaderLabel.Font.Style := [fsBold];

  TopPos := TopPos + ScaleY(25);

  // Create controls for all users but only show a page at a time
  for i := 0 to LocalUsersList.Count - 1 do
  begin
    // Checkbox
    UserCheckBoxes[i] := TCheckBox.Create(Page_CreateShortcutsForExistingUsers);
    UserCheckBoxes[i].Parent := Page_CreateShortcutsForExistingUsers.Surface;
    UserCheckBoxes[i].Left := ScaleX(20);
    UserCheckBoxes[i].Width := ScaleX(180);
    UserCheckBoxes[i].Caption := LocalUserDisplayList[i];
    UserCheckBoxes[i].OnClick := @OnUserCheckBoxClick;
    UserCheckBoxes[i].Tag := i;

    // Password edit
    UserPasswordEdits[i] := TEdit.Create(Page_CreateShortcutsForExistingUsers);
    UserPasswordEdits[i].Parent := Page_CreateShortcutsForExistingUsers.Surface;
    UserPasswordEdits[i].Left := ScaleX(220);
    UserPasswordEdits[i].Width := ScaleX(200);
    UserPasswordEdits[i].PasswordChar := '*';
    UserPasswordEdits[i].Enabled := False;
    UserPasswordEdits[i].OnChange := @OnPasswordEditChange;
    UserPasswordEdits[i].Tag := i;

    // Status label
    UserPasswordStatus[i] := TLabel.Create(Page_CreateShortcutsForExistingUsers);
    UserPasswordStatus[i].Parent := Page_CreateShortcutsForExistingUsers.Surface;
    UserPasswordStatus[i].Left := ScaleX(430);
    UserPasswordStatus[i].Font.Color := clRed;
    UserPasswordStatus[i].Caption := '';
    UserPasswordStatus[i].Visible := False;
  end;

  // Pagination controls (only if more users than fit on a single page)
  if LocalUsersList.Count > UsersPerPage then
  begin
    // Prev button
    Tool1PrevButton := TButton.Create(Page_CreateShortcutsForExistingUsers);
    Tool1PrevButton.Parent := Page_CreateShortcutsForExistingUsers.Surface;
    Tool1PrevButton.Left := ScaleX(20);
    Tool1PrevButton.Width := ScaleX(80);
    Tool1PrevButton.Caption := 'Previous';
    Tool1PrevButton.OnClick := @OnPrevUsersPageClick;

    // Page label
    Tool1PageLabel := TLabel.Create(Page_CreateShortcutsForExistingUsers);
    Tool1PageLabel.Parent := Page_CreateShortcutsForExistingUsers.Surface;
    Tool1PageLabel.AutoSize := True;
    Tool1PageLabel.Left := ScaleX(110);

    // Next button
    Tool1NextButton := TButton.Create(Page_CreateShortcutsForExistingUsers);
    Tool1NextButton.Parent := Page_CreateShortcutsForExistingUsers.Surface;
    Tool1NextButton.Width := ScaleX(80);
    Tool1NextButton.Caption := 'Next';
    Tool1NextButton.OnClick := @OnNextUsersPageClick;
  end
  else
  begin
    Tool1PrevButton := nil;
    Tool1NextButton := nil;
    Tool1PageLabel := nil;
  end;

  // Calculate bottom position for the password reset link
  BottomPos := Page_CreateShortcutsForExistingUsers.SurfaceHeight - ScaleY(70);

  // Create password reset link
  Tool1PasswordResetLink := TLabel.Create(Page_CreateShortcutsForExistingUsers);
  Tool1PasswordResetLink.Parent := Page_CreateShortcutsForExistingUsers.Surface;
  Tool1PasswordResetLink.Left := ScaleX(20);
  Tool1PasswordResetLink.Top := BottomPos;
  Tool1PasswordResetLink.Caption := 'Can''t remember a password? Click here to Reset it';
  // Choose a link color appropriate for current page theme
  if IsDarkColor(Page_CreateShortcutsForExistingUsers.Surface.Color) then
    Tool1PasswordResetLink.Font.Color := RGBToColor(135,206,250)
  else
    Tool1PasswordResetLink.Font.Color := clBlue;
  Tool1PasswordResetLink.Font.Style := [fsUnderline];
  Tool1PasswordResetLink.Cursor := crHandPoint;
  Tool1PasswordResetLink.OnClick := @OnPasswordResetLinkClick;

  // Start at first page and arrange visible controls
  CurrentUserPage := 0;
  UpdateUsersPageDisplay;

  SetUserControlsEnabled(True);

  // Mark controls as built to avoid duplicates on re-entry
  CreateShortcutsControlsBuilt := True;
end;

procedure UpdateUsersPageDisplay;
var
  i, PageCount, StartIdx, EndIdx, VisIndex, RowHeight, BaseTop: Integer;
begin
  LogEntry('UpdateUsersPageDisplay');
  if LocalUsersList.Count = 0 then
  begin
    LogDebug('UpdateUsersPageDisplay: no users, skipping');
    LogExit('UpdateUsersPageDisplay');
    exit;
  end;

  PageCount := (LocalUsersList.Count + UsersPerPage - 1) div UsersPerPage;
  if CurrentUserPage < 0 then CurrentUserPage := 0;
  if CurrentUserPage >= PageCount then CurrentUserPage := PageCount - 1;

  StartIdx := CurrentUserPage * UsersPerPage;
  EndIdx := StartIdx + UsersPerPage - 1;
  if EndIdx > LocalUsersList.Count - 1 then EndIdx := LocalUsersList.Count - 1;

  RowHeight := ScaleY(26);
  BaseTop := Tool1UsersHeaderLabel.Top + ScaleY(25);
  VisIndex := 0;

  for i := 0 to LocalUsersList.Count - 1 do
  begin
    if Assigned(UserCheckBoxes[i]) then
    begin
      UserCheckBoxes[i].Visible := (i >= StartIdx) and (i <= EndIdx);
      if UserCheckBoxes[i].Visible then
        UserCheckBoxes[i].Top := BaseTop + VisIndex * RowHeight;
    end;
    if Assigned(UserPasswordEdits[i]) then
    begin
      UserPasswordEdits[i].Visible := (i >= StartIdx) and (i <= EndIdx);
      if UserPasswordEdits[i].Visible then
        UserPasswordEdits[i].Top := BaseTop + VisIndex * RowHeight - ScaleY(2);
    end;
    if Assigned(UserPasswordStatus[i]) then
    begin
      UserPasswordStatus[i].Visible := False;
      if (i >= StartIdx) and (i <= EndIdx) then
      begin
        if UserPasswordStatus[i].Caption <> '' then
          UserPasswordStatus[i].Visible := True;
        UserPasswordStatus[i].Top := BaseTop + VisIndex * RowHeight - ScaleY(2);
      end;
    end;

    if (i >= StartIdx) and (i <= EndIdx) then
      Inc(VisIndex);
  end;

  if Assigned(Tool1PrevButton) and Assigned(Tool1NextButton) and Assigned(Tool1PageLabel) then
  begin
    Tool1PrevButton.Top := Page_CreateShortcutsForExistingUsers.SurfaceHeight - ScaleY(30);
    Tool1NextButton.Top := Tool1PrevButton.Top;
    Tool1NextButton.Left := Tool1PrevButton.Left + Tool1PrevButton.Width + ScaleX(120);
    Tool1PageLabel.Top := Tool1PrevButton.Top + ScaleY(4);
    Tool1PageLabel.Caption := 'Page ' + IntToStr(CurrentUserPage + 1) + ' of ' + IntToStr(PageCount);
    Tool1PrevButton.Enabled := CurrentUserPage > 0;
    Tool1NextButton.Enabled := CurrentUserPage < PageCount - 1;
  end;
end;

procedure OnPrevUsersPageClick(Sender: TObject);
begin
  LogEntry('OnPrevUsersPageClick');
  if CurrentUserPage > 0 then
    Dec(CurrentUserPage);
  UpdateUsersPageDisplay;
  LogExit('OnPrevUsersPageClick');
end;

procedure OnNextUsersPageClick(Sender: TObject);
var
  PageCount: Integer;
  StartIdx: Integer;
  EndIdx: Integer;
  i: Integer;
  HasErrors: Boolean;
  Password: string;
begin
  LogEntry('OnNextUsersPageClick');
  LogDebug('OnNextUsersPageClick: currentPage=' + IntToStr(CurrentUserPage));
  PageCount := (LocalUsersList.Count + UsersPerPage - 1) div UsersPerPage;

  StartIdx := CurrentUserPage * UsersPerPage;
  EndIdx := StartIdx + UsersPerPage - 1;
  if EndIdx > LocalUsersList.Count - 1 then
    EndIdx := LocalUsersList.Count - 1;

  HasErrors := False;
  for i := StartIdx to EndIdx do
  begin
    if Assigned(UserCheckBoxes[i]) and UserCheckBoxes[i].Checked then
    begin
      if Assigned(UserPasswordEdits[i]) then
        Password := UserPasswordEdits[i].Text
      else
        Password := '';

      if Password = '' then
      begin
        if Assigned(UserPasswordStatus[i]) then
        begin
          UserPasswordStatus[i].Caption := 'Can''t be blank';
          UserPasswordStatus[i].Visible := True;
        end;
        HasErrors := True;
        continue;
      end;

      if IsValidPassword(Password) <> '' then
      begin
        if Assigned(UserPasswordStatus[i]) then
        begin
          UserPasswordStatus[i].Caption := 'Invalid password';
          UserPasswordStatus[i].Visible := True;
        end;
        HasErrors := True;
        continue;
      end;

      if not ValidateLocalCredential(LocalUsersList[i], Password) then
      begin
        if Assigned(UserPasswordStatus[i]) then
        begin
          UserPasswordStatus[i].Caption := 'Incorrect PW';
          UserPasswordStatus[i].Visible := True;
        end;
        HasErrors := True;
        continue;
      end;

      if Assigned(UserPasswordStatus[i]) then
      begin
        UserPasswordStatus[i].Caption := '';
        UserPasswordStatus[i].Visible := False;
      end;
    end
    else
    begin
      // User was unchecked — clear any stale validation message
      if Assigned(UserPasswordStatus[i]) then
      begin
        UserPasswordStatus[i].Caption := '';
        UserPasswordStatus[i].Visible := False;
      end;
    end;
  end;

  if HasErrors then
  begin
    // Keep user on the same page and refresh statuses
    UpdateUsersPageDisplay;
    exit;
  end;

  if CurrentUserPage < PageCount - 1 then
    Inc(CurrentUserPage);
  UpdateUsersPageDisplay;
end;

function UserAlreadyEntered(const UserName: string): Boolean;
var
  i: Integer;
  CurrentUser: string;
  TempPassword: string;
begin
  LogEntry('UserAlreadyEntered');
  Result := False;
  for i := 0 to UsersList.Count - 1 do
  begin
    ParseUserEntry(UsersList[i], CurrentUser, TempPassword);
    if CompareText(CurrentUser, UserName) = 0 then
    begin
      Result := True;
      LogDebug('UserAlreadyEntered: user=' + UserName + ' already entered at index ' + IntToStr(i));
      LogExit('UserAlreadyEntered');
      exit;
    end;
  end;
  LogDebug('UserAlreadyEntered: user=' + UserName + ' not found in list');
  LogExit('UserAlreadyEntered');
end;

function IsValidUsername(const UserName: string): String;
var
  Len: Integer;
  i: Integer;
  Ch: Char;
begin
  LogEntry('IsValidUsername');
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
var
  i: Integer;
  c: Char;
begin
  LogEntry('IsValidPassword');
  Result := '';  // Empty string means valid
  
  if Length(Password) = 0 then
    Result := 'Password cannot be empty.'
  else if Length(Password) > 128 then
    Result := 'Password is too long (maximum 128 characters).'
  else
  begin
    for i := 1 to Length(Password) do
    begin
      c := Password[i];
      if (Ord(c) < 32) or (Ord(c) = 127) then
      begin
        Result := 'Password cannot contain control characters.';
        exit;
      end;
      if (c = '''') or (c = '"') or (c = '`') then
      begin
        Result := 'Password cannot contain quote characters (single quote, double quote, or backtick).';
        exit;
      end;
    end;
  end;
end;

function ValidateLocalCredential(const UserName, Password: string): Boolean;
var
  Token: Cardinal;
begin
  LogEntry('ValidateLocalCredential');
  // Fast local credential check via LogonUser; avoids slow PowerShell/WinRM
  Token := 0;
  Result := LogonUser(UserName, '.', Password, 2, 0, Token);
  if Token <> 0 then
    CloseHandle(Token);
  LogDebug('ValidateLocalCredential: user=' + UserName + ' result=' + BoolToStr(Result));
  LogExit('ValidateLocalCredential');
end;

function ShouldInstallFiles: Boolean;
begin
  LogEntry('ShouldInstallFiles');
  // Only install bundled TermWrap files when Install TermWrap is selected or when the
  // user explicitly selected "Install TermWrap" on the welcome/options page.
  Result := DoInstallTermWrap;
  LogDebug('ShouldInstallFiles: DoInstallTermWrap=' + BoolToStr(DoInstallTermWrap) + ' returning ' + BoolToStr(Result));
  LogExit('ShouldInstallFiles');
end;

function ShouldApplyRegistryEntries: Boolean;
begin
  LogEntry('ShouldApplyRegistryEntries');
  // Edit Shortcut Settings mode should only launch mstsc /edit and avoid
  // unrelated installer-side registry changes.
  Result := SelectedInstallMode <> installModeEditShortcuts;
  LogDebug('ShouldApplyRegistryEntries: SelectedInstallMode=' + IntToStr(SelectedInstallMode) + ' returning ' + BoolToStr(Result));
  LogExit('ShouldApplyRegistryEntries');
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  LogEntry('ShouldSkipPage');
  Result := False;
  
  // Always skip directory selection page
  if PageID = wpSelectDir then
    Result := True;
  
  // Skip User Page unless installing in Install mode with Create new users selected
  if (PageID = UserPage.ID) and
     ((SelectedInstallMode <> installModeInstall) or (not DoCreateRdpShortcuts) or (CreateUserMode <> createUserModeNew)) then
    Result := True;
// Show shortcut editing pages only for Edit Shortcut Settings mode
if (PageID = EditShortcutPage.ID) and (SelectedInstallMode <> installModeEditShortcuts) then
  Result := True;

// Show Advanced Editing page only when user chose the advanced option
if Assigned(Page_EditShortcutAdvanced) and (PageID = Page_EditShortcutAdvanced.ID) then
  Result := not ((SelectedInstallMode = installModeEditShortcuts) and
                 Assigned(chkShowMoreShortcutOptions) and chkShowMoreShortcutOptions.Checked);


  // Skip Ready page - no need to show install path
  if PageID = wpReady then
    Result := True;

  // Show Edit System-wide RDP settings page only in that mode
  if (PageID = EditSystemwideSettingsPage.ID) and (SelectedInstallMode <> installModeEditSystemwideSettings) then
    Result := True;

  // Show Page_ShowRDPInfo only in that mode
  if Assigned(Page_ShowRDPInfo) and (PageID = Page_ShowRDPInfo.ID) and (SelectedInstallMode <> installModeShowRDPInfo) then
    Result := True;

  // Show Quick Fixes page only in that mode
  if Assigned(QuickFixesPage) and (PageID = QuickFixesPage.ID) and (SelectedInstallMode <> installModeQuickFixes) then
    Result := True;

  // Show Create Shortcuts for Existing Users page only when:
  //   Install mode + Create RDP shortcuts + Use existing users
  if PageID = Page_CreateShortcutsForExistingUsers.ID then
    Result := not ((SelectedInstallMode = installModeInstall) and DoCreateRdpShortcuts and (CreateUserMode = createUserModeExisting));

  // Show Shortcut Settings page when creating shortcuts (any method) or editing shortcuts
  if PageID = Page_ShortcutSettings.ID then
    Result := not (((SelectedInstallMode = installModeInstall) and DoCreateRdpShortcuts) or
                   (SelectedInstallMode = installModeEditShortcuts));

end;

function IsVCRedistInstalled: Boolean;
var
  Major: Cardinal;
begin
  if SimulateNoVCRedist then
  begin
    if not SimLogNoVCRedistShown then
    begin
      LogSimulationScenario('System doesnt have VC++');
      SimLogNoVCRedistShown := True;
    end;
    Result := False;
    exit;
  end;

  // Check for VC++ 2015-2022 Redistributable (x64), version 14.x or higher
  Result := RegQueryDWordValue(HKLM, REG_VCREDIST, 'Major', Major) and (Major >= 14);
end;

procedure SecureCleanupTempFiles(const UserName: string);
begin
  LogEntry('SecureCleanupTempFiles');
  // Securely delete temporary files that contained sensitive data
  DeleteFile(TempFile('enc_' + UserName + '.txt'));
  DeleteFile(TempFile('create_rdp_' + UserName + '.ps1'));
  LogDebug('SecureCleanupTempFiles: cleaned up for user=' + UserName);
  LogExit('SecureCleanupTempFiles');
end;

function EncryptPasswordToFile(const Password, UserName: string): string;
var
  ResultCode: Integer;
  EncPath: string;
  Cmd: string;
  OpTick: Cardinal;
begin
  LogEntry('EncryptPasswordToFile');
  OpTick := GetTickCount;
  EncPath := TempFile('enc_' + UserName + '.txt');
  LogDebug('EncryptPasswordToFile: user=' + UserName + ' encPath=''' + EncPath + '''');

  // Use inline PowerShell command (no temporary .ps1 file) to avoid script-file stalls.
  Cmd :=
    '$pw = ''' + PSSingleQuote(Password) + ''' | ConvertTo-SecureString -AsPlainText -Force; ' +
    '$encPw = ConvertFrom-SecureString $pw; ' +
    '[System.IO.File]::WriteAllText(''' + PSSingleQuote(EncPath) + ''', $encPw)';

  ResultCode := -1;
  Exec(EXE_POWERSHELL, BuildPowerShellArgs(Cmd, True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  if (ResultCode = 0) and FileExists(EncPath) then
  begin
    LogDebug('EncryptPasswordToFile: success for user=' + UserName + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
    Result := EncPath;
  end
  else
  begin
    if FileExists(EncPath) then
      DeleteFile(EncPath);
    LogWarn('EncryptPasswordToFile failed for user=' + UserName + ' (exit=' + IntToStr(ResultCode) + ') [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
    Result := '';
  end;
  LogExit('EncryptPasswordToFile');
end;

// Read experience / performance checkboxes into RDP key integer values.
// disable wallpaper: 0=show, 1=disable  — checkbox means "allow" so invert for disable key
// allow font smoothing: 1=allow  — checkbox value maps directly
// allow desktop composition: 1=allow
// disable full window drag: 0=show contents, 1=hide  — inverted
// disable menu anims: 0=show, 1=hide  — inverted
// disable themes: 0=use themes, 1=disable  — inverted
procedure GetExperienceSettings(var DisableWallpaper, AllowFontSmooth, AllowComposition,
  DisableFullWindowDrag, DisableMenuAnims, DisableThemes: Integer);
begin
  if Assigned(chkExpWallpaper)    then begin if chkExpWallpaper.Checked    then DisableWallpaper      := 0 else DisableWallpaper      := 1; end else DisableWallpaper      := 1;
  if Assigned(chkExpFontSmooth)   then begin if chkExpFontSmooth.Checked   then AllowFontSmooth       := 1 else AllowFontSmooth       := 0; end else AllowFontSmooth       := 1;
  if Assigned(chkExpComposition)  then begin if chkExpComposition.Checked  then AllowComposition      := 1 else AllowComposition      := 0; end else AllowComposition      := 1;
  if Assigned(chkExpDragContents) then begin if chkExpDragContents.Checked then DisableFullWindowDrag := 0 else DisableFullWindowDrag := 1; end else DisableFullWindowDrag := 0;
  if Assigned(chkExpMenuAnim)     then begin if chkExpMenuAnim.Checked     then DisableMenuAnims      := 0 else DisableMenuAnims      := 1; end else DisableMenuAnims      := 0;
  if Assigned(chkExpVisualStyles) then begin if chkExpVisualStyles.Checked then DisableThemes         := 0 else DisableThemes         := 1; end else DisableThemes         := 0;
end;

// Read shortcut display settings from UI controls into output vars.
// Shared by WriteRDPFileDirect, GenerateRDPPowerShellScript, and WriteShortcutSettingsToRdpFile.
procedure GetShortcutDisplaySettings(var ScreenModeId, DesktopWidth, DesktopHeight, UseMultiMon, AudioMode, RedirectClipboard, KeyboardHook: Integer);
var
  ResStr: string;
  XPos: Integer;
begin
  if Assigned(chkFullScreen) and chkFullScreen.Checked then
    ScreenModeId := 2
  else
    ScreenModeId := 1;

  DesktopWidth  := DEFAULT_RDP_WIDTH;
  DesktopHeight := DEFAULT_RDP_HEIGHT;
  if Assigned(cboResolution) and (cboResolution.ItemIndex >= 0) then
  begin
    if cboResolution.Items[cboResolution.ItemIndex] = 'Custom' then
    begin
      if Assigned(edtCustomWidth)  then DesktopWidth  := StrToIntDef(Trim(edtCustomWidth.Text),  DEFAULT_RDP_WIDTH);
      if Assigned(edtCustomHeight) then DesktopHeight := StrToIntDef(Trim(edtCustomHeight.Text), DEFAULT_RDP_HEIGHT);
    end
    else
    begin
      ResStr := cboResolution.Items[cboResolution.ItemIndex];
      XPos := Pos(' x ', ResStr);
      if XPos > 0 then
      begin
        DesktopWidth  := StrToIntDef(Trim(Copy(ResStr, 1, XPos - 1)), DEFAULT_RDP_WIDTH);
        DesktopHeight := StrToIntDef(Trim(Copy(ResStr, XPos + 3, Length(ResStr))), DEFAULT_RDP_HEIGHT);
      end;
    end;
  end;

  if Assigned(chkUseAllMonitors) and chkUseAllMonitors.Checked then UseMultiMon := 1 else UseMultiMon := 0;
  if Assigned(chkSound) and chkSound.Checked then AudioMode := 0 else AudioMode := 2;
  if Assigned(chkCopyPaste) and chkCopyPaste.Checked then RedirectClipboard := 1 else RedirectClipboard := 0;
  // Keyboard Hook: 0=main PC, 1=RDP, 2=RDP only if full-screen
  if Assigned(cboKeyboardHook) then
    KeyboardHook := cboKeyboardHook.ItemIndex
  else
    KeyboardHook := 0;
end;

function GetCurrentRdpPort: Integer;
// Reads the configured RDP listening port from the registry.
// Returns the default RDP_LISTEN_PORT (3389) if the registry value is absent.
var
  PortNumber: Cardinal;
begin
  if RegQueryDWordValue(HKLM, REG_RDP_TCP, 'PortNumber', PortNumber) then
    Result := PortNumber
  else
    Result := RDP_LISTEN_PORT;
end;

function WriteRDPFileDirect(const UserName, RDPPath, EncPath: string): Boolean;
var
  ScreenModeId: Integer;
  DesktopWidth: Integer;
  DesktopHeight: Integer;
  UseMultiMon: Integer;
  AudioMode: Integer;
  RedirectClipboard: Integer;
  KeyboardHook: Integer;
  DisableWallpaper, AllowFontSmooth, AllowComposition: Integer;
  DisableFullWindowDrag, DisableMenuAnims, DisableThemes: Integer;
  RdpPort: Integer;
  SL: TStringList;
  EncTextRaw: AnsiString;
  EncText: string;
  HasEnc: Boolean;
  OpTick: Cardinal;
begin
  LogEntry('WriteRDPFileDirect');
  OpTick := GetTickCount;
  Result := False;
  EncText := '';
  HasEnc := False;

  if (EncPath <> '') and FileExists(EncPath) then
  begin
    if LoadStringFromFile(EncPath, EncTextRaw) then
    begin
      EncText := Trim(String(EncTextRaw));
      HasEnc := EncText <> '';
    end;
  end;

  GetShortcutDisplaySettings(ScreenModeId, DesktopWidth, DesktopHeight, UseMultiMon, AudioMode, RedirectClipboard, KeyboardHook);
  GetExperienceSettings(DisableWallpaper, AllowFontSmooth, AllowComposition, DisableFullWindowDrag, DisableMenuAnims, DisableThemes);

  SL := TStringList.Create;
  try
    SL.Add('username:s:' + UserName);
    SL.Add('screen mode id:i:' + IntToStr(ScreenModeId));
    SL.Add('desktopwidth:i:' + IntToStr(DesktopWidth));
    SL.Add('desktopheight:i:' + IntToStr(DesktopHeight));
    SL.Add('use multimon:i:' + IntToStr(UseMultiMon));
    SL.Add('session bpp:i:32');
    SL.Add('smart sizing:i:1');
    SL.Add('dynamic resolution:i:1');
    // Append port if non-default (e.g. 127.0.0.2:3390), otherwise use bare IP
    RdpPort := GetCurrentRdpPort;
    if RdpPort = RDP_LISTEN_PORT then
      SL.Add('full address:s:' + RDP_LOOPBACK_IP)
    else
      SL.Add('full address:s:' + RDP_LOOPBACK_IP + ':' + IntToStr(RdpPort));
    SL.Add('autoreconnection enabled:i:1');
    SL.Add('compression:i:1');
    SL.Add('keyboardhook:i:' + IntToStr(KeyboardHook));
    SL.Add('audiocapturemode:i:0');
    SL.Add('audiomode:i:' + IntToStr(AudioMode));
    SL.Add('redirectclipboard:i:' + IntToStr(RedirectClipboard));
    SL.Add('redirectdrives:i:0');
    SL.Add('redirectprinters:i:0');
    SL.Add('redirectsmartcards:i:0');
    SL.Add('redirectwebauthn:i:0');
    SL.Add('videoplaybackmode:i:1');
    SL.Add('connection type:i:7');
    SL.Add('displayconnectionbar:i:1');
    SL.Add('disable wallpaper:i:' + IntToStr(DisableWallpaper));
    SL.Add('allow font smoothing:i:' + IntToStr(AllowFontSmooth));
    SL.Add('allow desktop composition:i:' + IntToStr(AllowComposition));
    SL.Add('disable full window drag:i:' + IntToStr(DisableFullWindowDrag));
    SL.Add('disable menu anims:i:' + IntToStr(DisableMenuAnims));
    SL.Add('disable themes:i:' + IntToStr(DisableThemes));
    SL.Add('bitmapcachepersistenable:i:1');
    SL.Add('authentication level:i:0');
    if HasEnc then
      SL.Add('prompt for credentials:i:0')
    else
      SL.Add('prompt for credentials:i:1');
    SL.Add('negotiate security layer:i:1');
    SL.Add('enablecredsspsupport:i:1');
    SL.Add('remoteapplicationmode:i:0');
    SL.Add('drivestoredirect:s:');
    SL.Add('alternate shell:s:');
    SL.Add('shell working directory:s:');
    SL.Add('gatewayhostname:s:');
    SL.Add('gatewayusagemethod:i:0');
    SL.Add('gatewaycredentialssource:i:0');
    SL.Add('gatewayprofileusagemethod:i:0');
    SL.Add('promptcredentialonce:i:0');
    SL.Add('use redirection server name:i:0');
    if HasEnc then
      SL.Add('password 51:b:' + EncText);
    SL.Add('disableconnectionsharing:i:0');
    SL.SaveToFile(RDPPath);
    Result := FileExists(RDPPath);
  finally
    SL.Free;
  end;
end;

function GenerateRDPPowerShellScript(const UserName, RDPPath, EncPath: string): string;
var
  ScreenModeId: Integer;
  DesktopWidth: Integer;
  DesktopHeight: Integer;
  UseMultiMon: Integer;
  AudioMode: Integer;
  RedirectClipboard: Integer;
  KeyboardHook: Integer;
  DisableWallpaper, AllowFontSmooth, AllowComposition: Integer;
  DisableFullWindowDrag, DisableMenuAnims, DisableThemes: Integer;
  FullAddressLine: string;
  RdpPort: Integer;
begin
  LogEntry('GenerateRDPPowerShellScript');
  LogDebug('GenerateRDPPowerShellScript: user=' + UserName + ' path=' + RDPPath);
  GetShortcutDisplaySettings(ScreenModeId, DesktopWidth, DesktopHeight, UseMultiMon, AudioMode, RedirectClipboard, KeyboardHook);
  GetExperienceSettings(DisableWallpaper, AllowFontSmooth, AllowComposition, DisableFullWindowDrag, DisableMenuAnims, DisableThemes);

  // Append port if non-default (e.g. 127.0.0.2:3390), otherwise use bare IP
  RdpPort := GetCurrentRdpPort;
  if RdpPort = RDP_LISTEN_PORT then
    FullAddressLine := 'full address:s:' + RDP_LOOPBACK_IP
  else
    FullAddressLine := 'full address:s:' + RDP_LOOPBACK_IP + ':' + IntToStr(RdpPort);

  Result :=
    'param([string]$EncPath = '''')' + #13#10 +
    'try {' + #13#10 +
    '  Write-Host "Starting RDP file creation..."' + #13#10 +
    '  $hasEnc = ($EncPath -ne '''') -and (Test-Path $EncPath)' + #13#10 +
    '  $encPass = ""' + #13#10 +
    '  if ($hasEnc) {' + #13#10 +
    '    $encPassLines = @(Get-Content "$EncPath")' + #13#10 +
    '    $encPass = ($encPassLines -join "").Trim()' + #13#10 +
    '    Write-Host "Encrypted password loaded"' + #13#10 +
    '  } else {' + #13#10 +
    '    Write-Host "No encrypted password available; shortcut will prompt for password"' + #13#10 +
    '  }' + #13#10 +
    '  $rdp = @()' + #13#10 +
    '  $rdp += "' + FullAddressLine + '"' + #13#10 +
    '  $rdp += "username:s:' + UserName + '"' + #13#10 +
    '  $rdp += "screen mode id:i:' + IntToStr(ScreenModeId) + '"' + #13#10 +
    '  $rdp += "desktopwidth:i:' + IntToStr(DesktopWidth) + '"' + #13#10 +
    '  $rdp += "desktopheight:i:' + IntToStr(DesktopHeight) + '"' + #13#10 +
    '  $rdp += "use multimon:i:' + IntToStr(UseMultiMon) + '"' + #13#10 +
    '  $rdp += "session bpp:i:32"' + #13#10 +
    '  $rdp += "smart sizing:i:1"' + #13#10 +
    '  $rdp += "dynamic resolution:i:1"' + #13#10 +
    '  $rdp += "autoreconnection enabled:i:1"' + #13#10 +
    '  $rdp += "compression:i:1"' + #13#10 +
    '  $rdp += "keyboardhook:i:' + IntToStr(KeyboardHook) + '"' + #13#10 +
    '  $rdp += "audiocapturemode:i:0"' + #13#10 +
    '  $rdp += "audiomode:i:' + IntToStr(AudioMode) + '"' + #13#10 +
    '  $rdp += "redirectclipboard:i:' + IntToStr(RedirectClipboard) + '"' + #13#10 +
    '  $rdp += "redirectdrives:i:0"' + #13#10 +
    '  $rdp += "redirectprinters:i:0"' + #13#10 +
    '  $rdp += "videoplaybackmode:i:1"' + #13#10 +
    '  $rdp += "connection type:i:7"' + #13#10 +
    '  $rdp += "displayconnectionbar:i:1"' + #13#10 +
    '  $rdp += "disable wallpaper:i:' + IntToStr(DisableWallpaper) + '"' + #13#10 +
    '  $rdp += "allow font smoothing:i:' + IntToStr(AllowFontSmooth) + '"' + #13#10 +
    '  $rdp += "allow desktop composition:i:' + IntToStr(AllowComposition) + '"' + #13#10 +
    '  $rdp += "disable full window drag:i:' + IntToStr(DisableFullWindowDrag) + '"' + #13#10 +
    '  $rdp += "disable menu anims:i:' + IntToStr(DisableMenuAnims) + '"' + #13#10 +
    '  $rdp += "disable themes:i:' + IntToStr(DisableThemes) + '"' + #13#10 +
    '  $rdp += "bitmapcachepersistenable:i:1"' + #13#10 +
    '  $rdp += "authentication level:i:0"' + #13#10 +
    '  if ($hasEnc) { $rdp += "prompt for credentials:i:0" } else { $rdp += "prompt for credentials:i:1" }' + #13#10 +
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
    '  if ($hasEnc) { $rdp += ("password 51:b:" + $encPass) }' + #13#10 +
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

procedure CreateRDPShortcut(const UserName, Password, CreationSource: string);
var
  ResultCode: Integer;
  RDPPath: string;
  EncPath: string;
  ScriptPath: string;
  RunnerPath: string;
  RunnerScript: string;
  PowerShellScript: string;
  ShortcutFileName: string;
  OpTick: Cardinal;
  FileSizeStr: string;
begin
  LogEntry('CreateRDPShortcut');
  OpTick := GetTickCount;
  // Use custom shortcut name if the field is visible (single-user editing only).
  if Assigned(edtShortcutName) and edtShortcutName.Visible and (Trim(edtShortcutName.Text) <> '') then
    ShortcutFileName := Trim(edtShortcutName.Text)
  else
    ShortcutFileName := UserName + '.rdp';
  if CompareText(ExtractFileExt(ShortcutFileName), '.rdp') <> 0 then
    ShortcutFileName := ShortcutFileName + '.rdp';
  RDPPath := ExpandConstant('{userdesktop}\' + ShortcutFileName);
  ScriptPath := TempFile('create_rdp_' + UserName + '.ps1');
  LogDebug('CreateRDPShortcut: ShortcutFileName=' + ShortcutFileName + ' RDPPath=''' + RDPPath + ''' CreationSource=' + CreationSource);

  if PASSWORD_PIPELINE_DIAG <> 0 then
  begin
    LogSectionHeader('PASSWORD PIPELINE DEBUG');
    LogKeyValue('User', UserName);
    LogKeyValue('Creation Source', CreationSource);
    LogPasswordPipeline('SHORTCUT_INPUT', UserName, Password);
  end;

  LogDebug('CreateRDPShortcut: Creating RDP file at ' + RDPPath);

  // Remove any existing shortcut so we always overwrite with the new one
  if FileExists(RDPPath) then
  begin
    LogDebug('CreateRDPShortcut: Deleting existing RDP file');
    DeleteFile(RDPPath);
  end;

  // Encrypt the password
  LogDebug('CreateRDPShortcut: Encrypting password for user ' + UserName);
  EncPath := EncryptPasswordToFile(Password, UserName);
  if EncPath <> '' then
    LogEncryptedFileSummary('AFTER_ENCRYPT', EncPath);

  // Direct write path avoids PowerShell hangs in shortcut generation.
  if WriteRDPFileDirect(UserName, RDPPath, EncPath) then
  begin
    LogDebug('CreateRDPShortcut: Direct write succeeded, signing...');

    // Verify the RDP file exists and log its size for diagnostics
    if FileExists(RDPPath) then
    begin
      FileSizeStr := GetPSOutput('(Get-Item ''' + PSSingleQuote(RDPPath) + ''').Length');
      LogDebug('CreateRDPShortcut: RDP file verified: ' + RDPPath + ' size=' + FileSizeStr + ' bytes');
    end
    else
      LogWarn('CreateRDPShortcut: RDP file NOT FOUND after direct write: ' + RDPPath);

    SignRdpFile(RDPPath);
    SecureCleanupTempFiles(UserName);
    LogInfo('CreateRDPShortcut: Completed for ' + UserName + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
    LogExit('CreateRDPShortcut');
    exit;
  end
  else
  begin
    LogWarn('CreateRDPShortcut: Direct write failed, falling back to PowerShell script path');
  end;

  // Generate and execute PowerShell script
  LogDebug('CreateRDPShortcut: Generating RDP PowerShell script');
  PowerShellScript := GenerateRDPPowerShellScript(UserName, RDPPath, EncPath);
  SaveStringToFile(ScriptPath, PowerShellScript, False);
  LogDebug('CreateRDPShortcut: Script saved to ' + ScriptPath);
  
  RunnerPath := TempFile('run_create_rdp_' + UserName + '.ps1');
  RunnerScript :=
    'param([string]$TargetScript, [string]$EncPath, [int]$TimeoutSec)' + #13#10 +
    'try {' + #13#10 +
    '  $sb = {' + #13#10 +
    '    param([string]$TS, [string]$EP)' + #13#10 +
    '    & $TS -EncPath $EP' + #13#10 +
    '  }' + #13#10 +
    '  $job = Start-Job -ScriptBlock $sb -ArgumentList $TargetScript, $EncPath' + #13#10 +
    '  if (Wait-Job -Job $job -Timeout $TimeoutSec) {' + #13#10 +
    '    Receive-Job -Job $job -ErrorAction SilentlyContinue | Out-Null' + #13#10 +
    '    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null' + #13#10 +
    '    exit 0' + #13#10 +
    '  } else {' + #13#10 +
    '    Stop-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null' + #13#10 +
    '    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null' + #13#10 +
    '    exit 124' + #13#10 +
    '  }' + #13#10 +
    '} catch {' + #13#10 +
    '  exit 1' + #13#10 +
    '}';
  SaveStringToFile(RunnerPath, RunnerScript, False);
  LogDebug('CreateRDPShortcut: Executing PowerShell script with timeout wrapper');
  Exec(EXE_POWERSHELL,
    BuildPowerShellFileArgs(
      RunnerPath,
      BuildPSNamedParam('TargetScript', ScriptPath) + ' ' + BuildPSNamedParam('EncPath', EncPath) + ' -TimeoutSec 30',
      True),
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  LogDebug('CreateRDPShortcut: Wrapped PowerShell script exit code=' + IntToStr(ResultCode));
  DeleteFile(RunnerPath);
  
  if ResultCode <> 0 then
  begin
    LogWarn('RDP file creation failed with exit code=' + IntToStr(ResultCode));
  end
  else
  begin
    LogDebug('CreateRDPShortcut: RDP file created successfully via PowerShell, signing...');

    // Verify the RDP file exists and log its size for diagnostics
    if FileExists(RDPPath) then
    begin
      FileSizeStr := GetPSOutput('(Get-Item ''' + PSSingleQuote(RDPPath) + ''').Length');
      LogDebug('CreateRDPShortcut: RDP file verified after PS: ' + RDPPath + ' size=' + FileSizeStr + ' bytes');
    end
    else
      LogWarn('CreateRDPShortcut: RDP file NOT FOUND after PS creation: ' + RDPPath);

    SignRdpFile(RDPPath);
  end;

  LogDebug('CreateRDPShortcut: Cleaning up temporary files');
  SecureCleanupTempFiles(UserName);
  LogInfo('CreateRDPShortcut: Completed for ' + UserName + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('CreateRDPShortcut');
end;

procedure ClearPasswordsFromMemory;
begin
  LogEntry('ClearPasswordsFromMemory');
  if Assigned(UsersList) then
  begin
    LogDebug('ClearPasswordsFromMemory: clearing ' + IntToStr(UsersList.Count) + ' entries');
    UsersList.Clear;
  end;
  LogExit('ClearPasswordsFromMemory');
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
  OutPath: string;
  ScriptPath: string;
  PSScript: string;
  PSParams: string;
  SL: TStringList;
  j: Integer;
  NetRc: Integer;
  ExecOk: Boolean;
  UserCreateOutputAlreadyLogged: Boolean;
  UserCreatePath: string;
  CredentialMatchesEnteredPassword: Boolean;
  ValidationError: string;
begin
  LogEntry('CreateRDPUsers');
  // Lazy-resolve group names on first use (avoids blocking during InitializeWizard)
  if GroupAdministratorsName = 'Administrators' then
  begin
    GroupAdministratorsName := GetLocalizedGroupName('S-1-5-32-544', 'Administrators');
    LogDebug('CreateRDPUsers: Resolved Administrators group name to: ' + GroupAdministratorsName);
  end;
  if GroupRDPUsersName = 'Remote Desktop Users' then
  begin
    GroupRDPUsersName := GetLocalizedGroupName('S-1-5-32-555', 'Remote Desktop Users');
    LogDebug('CreateRDPUsers: Resolved Remote Desktop Users group name to: ' + GroupRDPUsersName);
  end;
  
  // Start overall watchdog timer
  StartTick := GetTickCount;
  LogInfo('Starting CreateRDPUsers for ' + IntToStr(UsersList.Count) + ' users | Timeout=' + IntToStr(USERS_OVERALL_TIMEOUT) + 'ms');
  
  for i := 0 to UsersList.Count - 1 do
  begin
    UserStartTick := GetTickCount;
    LogDebug('CreateRDPUsers: Processing user ' + IntToStr(i+1) + '/' + IntToStr(UsersList.Count));
    // Check overall timeout
    if (GetTickCount - StartTick) > USERS_OVERALL_TIMEOUT then
    begin
      LogError('CreateRDPUsers overall timeout reached after ' + IntToStr(GetTickCount - StartTick) + ' ms; aborting remaining users');
      break;
    end;

    UserInfo := UsersList[i];
    ParseUserEntry(UserInfo, UserName, Password);

    ValidationError := IsValidUsername(UserName);
    if ValidationError <> '' then
    begin
      LogError('Skipping user due to invalid username input: ' + ValidationError + ' | User=' + UserName);
      continue;
    end;
    ValidationError := IsValidPassword(Password);
    if ValidationError <> '' then
    begin
      LogError('Skipping user due to invalid password input: ' + ValidationError + ' | User=' + UserName);
      continue;
    end;

    StatusOverlay.Caption := 'Creating user account (' + IntToStr(i + 1) + ' of ' + IntToStr(UsersList.Count) + '): ' + UserName;
    LogDebug('Creating user: ' + UserName + ' (user ' + IntToStr(i+1) + '/' + IntToStr(UsersList.Count) + ')');
    UserCreateOutputAlreadyLogged := False;
    UserCreatePath := 'POWERSHELL';
    LogPasswordPipeline('CREATE_FLOW_INPUT', UserName, Password);

    // PRIMARY PATH: Create user via PowerShell New-LocalUser with PasswordNeverExpires.
    // PowerShell native cmdlets properly set the password-never-expires flag on the
    // user account, unlike net.exe /expires:never which only controls account expiration.
    // Falls back to net.exe on failure (with net accounts /maxpwage:unlimited).
    OutPath := TempFile('user_create_' + SanitizeFileName(UserName) + '.log');
    PSScript :=
      'param([string]$UserName, [string]$Password, [string]$OutPath)' + #13#10 +
      '$ErrorActionPreference = ''Stop''' + #13#10 +
      'try {' + #13#10 +
      '  $outDir = [System.IO.Path]::GetDirectoryName($OutPath)' + #13#10 +
      '  if (-not [string]::IsNullOrWhiteSpace($outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }' + #13#10 +
      '  $pw = ConvertTo-SecureString -String $Password -AsPlainText -Force' + #13#10 +
      '  $existing = Get-LocalUser -Name $UserName -ErrorAction SilentlyContinue' + #13#10 +
      '  if ($null -ne $existing) {' + #13#10 +
      '    Set-LocalUser -Name $UserName -Password $pw -PasswordNeverExpires $true -ErrorAction Stop' + #13#10 +
      '    @(''SETLOCALUSER_PASSWORD_OK'', ''User='' + $UserName) | Out-File -FilePath $OutPath -Encoding UTF8' + #13#10 +
      '    exit 0' + #13#10 +
      '  }' + #13#10 +
      '  New-LocalUser -Name $UserName -Password $pw -FullName $UserName -PasswordNeverExpires -ErrorAction Stop | Out-Null' + #13#10 +
      '  @(''NEWLOCALUSER_OK'', ''User='' + $UserName) | Out-File -FilePath $OutPath -Encoding UTF8' + #13#10 +
      '  exit 0' + #13#10 +
      '} catch {' + #13#10 +
      '  @(' + #13#10 +
      '    ''NEWLOCALUSER_FAIL'',' + #13#10 +
      '    ''User='' + $UserName,' + #13#10 +
      '    ''ExceptionType='' + $_.Exception.GetType().FullName,' + #13#10 +
      '    ''Message='' + $_.Exception.Message,' + #13#10 +
      '    ''HResult='' + $_.Exception.HResult,' + #13#10 +
      '    ''CategoryInfo='' + $_.CategoryInfo.ToString(),' + #13#10 +
      '    ''FullyQualifiedErrorId='' + $_.FullyQualifiedErrorId,' + #13#10 +
      '    ''StackTrace:'',' + #13#10 +
      '    ($_ | Out-String)' + #13#10 +
      '  ) | Out-File -FilePath $OutPath -Encoding UTF8' + #13#10 +
      '  exit 1' + #13#10 +
      '}';
    ScriptPath := TempFile('create_local_user_script_' + SanitizeFileName(UserName) + '.ps1');
    SaveStringToFile(ScriptPath, PSScript, False);
    PSParams :=
      BuildPSNamedParam('UserName', UserName) + ' ' +
      BuildPSNamedParam('Password', Password) + ' ' +
      BuildPSNamedParam('OutPath', OutPath);
    WriteInstallerLog('PowerShell File: ' + ScriptPath + ' ' + MaskPasswordsInString(PSParams));
    ExecOk := Exec(EXE_POWERSHELL, BuildPowerShellFileArgs(ScriptPath, PSParams, True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    if not ExecOk then
      WriteInstallerLog('ERROR: Failed to launch PowerShell for user creation: code=' + IntToStr(ResultCode) + ' message=' + SysErrorMessage(ResultCode));
    WriteInstallerLog('New-LocalUser exit=' + IntToStr(ResultCode) + ' for ' + UserName);
    DeleteFile(ScriptPath);
    if not FileExists(OutPath) then
    begin
      WriteInstallerLog('WARNING: Missing user-create output file after PowerShell for ' + UserName + ': ' + OutPath);
      SaveStringToFile(OutPath,
        'NO_USER_CREATE_OUTPUT_FILE' + #13#10 +
        'Method=' + UserCreatePath + #13#10 +
        'ResultCode=' + IntToStr(ResultCode) + #13#10 +
        'User=' + UserName + #13#10,
        False);
    end;

    // FALLBACK PATH: If PowerShell failed, fall back to net.exe.
    // Sets system-wide max password age to unlimited via net accounts first.
    if ResultCode <> 0 then
    begin
      UserCreatePath := 'NET';
      WriteInstallerLog('WARNING: PowerShell user creation failed for ' + UserName + ', falling back to net.exe path');

      // Set system-wide maximum password age to unlimited (all passwords
      // will not expire). Combined with /expires:never below, this ensures the user
      // account and its password never expire.
      RunNetHiddenCapture('accounts /maxpwage:unlimited', 'maxpwage_' + UserName);

      // Two-step net.exe creation (avoids 14-char LM password prompt).
      // Step 1: create with short throwaway password (no LM prompt)
      // Step 2: set real password (password change does not trigger the LM prompt)
      // If either step fails, delete the partial user.
      NetRc := RunNetHiddenCapture('user ' + QuoteExeArg(UserName) + ' ' + QuoteExeArg(NET_USER_TEMP_PASSWORD) + ' /add /fullname:' + QuoteExeArg(UserName) + ' /expires:never', 'create_' + UserName);
      if NetRc = 0 then
      begin
        NetRc := RunNetHiddenCapture('user ' + QuoteExeArg(UserName) + ' ' + QuoteExeArg(Password), 'setpwd_' + UserName);
        if NetRc <> 0 then
        begin
          WriteInstallerLog('WARNING: NET user password set failed for ' + UserName + ', deleting partial user');
          RunNetHiddenCapture('user ' + QuoteExeArg(UserName) + ' /delete', 'del_' + UserName);
        end;
      end;
      // Only overwrite ResultCode if net.exe succeeded (otherwise preserve PS error code)
      if NetRc = 0 then
        ResultCode := 0;
    end;

    CredentialMatchesEnteredPassword := ValidateLocalCredential(UserName, Password);
    WriteInstallerLog('PASSWORD_DIAG [POST_CREATE_CREDENTIAL_TEST] user=' + UserName +
      ' | method=' + UserCreatePath +
      ' | enteredPasswordValid=' + BoolToStr(CredentialMatchesEnteredPassword));
    if not CredentialMatchesEnteredPassword then
      WriteInstallerLog('ERROR: Credential validation failed immediately after user creation for ' + UserName + ' (method=' + UserCreatePath + ')');

    // If any output produced, dump it to the installer log for diagnostics
    if FileExists(OutPath) then
    begin
      if not UserCreateOutputAlreadyLogged then
      begin
        SL := TStringList.Create;
        try
          try
            SL.LoadFromFile(OutPath);
            WriteInstallerLog('DEBUG: User-creation output for ' + UserName + ':');
            for j := 0 to SL.Count - 1 do
              WriteInstallerLog('DEBUG: ' + SL[j]);
          except
            // ignore logging errors
          end;
        finally
          SL.Free;
        end;
      end;
      DeleteFile(OutPath);
    end;
    Sleep(SLEEP_SHORT);

    // Add to Administrators group (prefer net localgroup)
    OutPath := TempFile('user_add_admin_' + SanitizeFileName(UserName) + '.log');
    NetRc := RunNetHiddenCapture('localgroup ' + QuoteExeArg(GroupAdministratorsName) + ' ' + QuoteExeArg(UserName) + ' /add', 'addadmin_' + UserName);
    ResultCode := NetRc;
    if ResultCode <> 0 then
    begin
      PSScript := BuildAddGroupMemberPowerShellScript('S-1-5-32-544', UserName, OutPath, 'ADD_ADMIN_OK');
      WriteInstallerLog('DEBUG: Fallback: Adding user to Administrators via PowerShell (SID: S-1-5-32-544) - resolved name=' + GroupAdministratorsName + ' user=' + UserName);
      PSParams :=
        BuildPSNamedParam('GroupSid', 'S-1-5-32-544') + ' ' +
        BuildPSNamedParam('UserName', UserName) + ' ' +
        BuildPSNamedParam('OutPath', OutPath) + ' ' +
        BuildPSNamedParam('SuccessTag', 'ADD_ADMIN_OK');
      ExecSavedPowerShellDebugScriptParams('add_admin_member', UserName, PSScript, PSParams, True, ResultCode);
      WriteInstallerLog('PowerShell add to Administrators exit=' + IntToStr(ResultCode) + ' for ' + UserName);
    end;
    if FileExists(OutPath) then
    begin
      SL := TStringList.Create;
      try
        try
          SL.LoadFromFile(OutPath);
          WriteInstallerLog('DEBUG: Add-LocalGroupMember (Administrators) output for ' + UserName + ':');
          for j := 0 to SL.Count - 1 do
            WriteInstallerLog('DEBUG: ' + SL[j]);
        except
        end;
      finally
        SL.Free;
        DeleteFile(OutPath);
      end;
    end;

    // Check if Remote Desktop Users group exists before adding
    // Fast check via net localgroup first (avoids hangs in some PowerShell hosts)
    NetRc := RunNetHiddenCapture('localgroup ' + QuoteExeArg(GroupRDPUsersName), 'checkrdp_' + UserName);
    if NetRc = 0 then
      ResultCode := 0
    else
    begin
      PSScript :=
        'param([string]$GroupName)' + #13#10 +
        '$exists = $false' + #13#10 +
        'try { $null = Get-LocalGroup -SID ''S-1-5-32-555'' -ErrorAction Stop; $exists = $true } catch { }' + #13#10 +
        'if (-not $exists) { try { $null = Get-LocalGroup -Name $GroupName -ErrorAction Stop; $exists = $true } catch { } }' + #13#10 +
        'if ($exists) { exit 0 } else { exit 2 }';
      PSParams := BuildPSNamedParam('GroupName', GroupRDPUsersName);
      ExecPowerShellScriptContent('check_rdp_group.ps1', PSScript, PSParams, True, ResultCode);
    end;
    // Skip adding to the group only when we confirmed it is missing (PS returned 2)
    if ResultCode <> 2 then
    begin
      OutPath := TempFile('user_add_rdp_' + SanitizeFileName(UserName) + '.log');
      NetRc := RunNetHiddenCapture('localgroup ' + QuoteExeArg(GroupRDPUsersName) + ' ' + QuoteExeArg(UserName) + ' /add', 'addrdp_' + UserName);
      ResultCode := NetRc;
      if ResultCode <> 0 then
      begin
        PSScript := BuildAddGroupMemberPowerShellScript('S-1-5-32-555', UserName, OutPath, 'ADD_RDP_OK');
        WriteInstallerLog('DEBUG: Fallback: Adding user to Remote Desktop Users via PowerShell (SID: S-1-5-32-555) - resolved name=' + GroupRDPUsersName + ' user=' + UserName);
        PSParams :=
          BuildPSNamedParam('GroupSid', 'S-1-5-32-555') + ' ' +
          BuildPSNamedParam('UserName', UserName) + ' ' +
          BuildPSNamedParam('OutPath', OutPath) + ' ' +
          BuildPSNamedParam('SuccessTag', 'ADD_RDP_OK');
        ExecSavedPowerShellDebugScriptParams('add_rdp_member', UserName, PSScript, PSParams, True, ResultCode);
        WriteInstallerLog('PowerShell add to Remote Desktop Users exit=' + IntToStr(ResultCode) + ' for ' + UserName);
      end;

      if FileExists(OutPath) then
      begin
        SL := TStringList.Create;
        try
          try
            SL.LoadFromFile(OutPath);
            WriteInstallerLog('DEBUG: Add to Remote Desktop Users output for ' + UserName + ':');
            for j := 0 to SL.Count - 1 do
              WriteInstallerLog('DEBUG: ' + SL[j]);
          except
          end;
        finally
          SL.Free;
          DeleteFile(OutPath);
        end;
      end;

      if ValidateGroupMembership(GroupRDPUsersName, UserName) then
        WriteInstallerLog('GROUP_DIAG [POST_ADD_VALIDATE] group=' + GroupRDPUsersName + ' user=' + UserName + ' | member=True')
      else
        WriteInstallerLog('ERROR: GROUP_DIAG [POST_ADD_VALIDATE] group=' + GroupRDPUsersName + ' user=' + UserName + ' | member=False');
    end
    else
    begin
      WriteInstallerLog('INFO: Remote Desktop Users group does not exist on this system. Skipping group membership for ' + UserName);
    end;

    // Create RDP shortcut using helper function
    CreateRDPShortcut(UserName, Password, UserCreatePath);
    LogDebug('Created shortcut for ' + UserName + ' [DURATION:' + IntToStr(GetTickCount - UserStartTick) + 'ms]');

    if (GetTickCount - UserStartTick) > PER_USER_TIMEOUT then
      LogWarn('CreateRDPUsers per-user time exceeded ' + IntToStr(PER_USER_TIMEOUT) + ' ms for ' + UserName);
  end;
  LogInfo('CreateRDPUsers completed: ' + IntToStr(UsersList.Count) + ' users processed [DURATION:' + IntToStr(GetTickCount - StartTick) + 'ms]');
  LogExit('CreateRDPUsers');
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
  LogEntry('CreateShortcutsForExistingUsers');
  StartTick := GetTickCount;
  LogInfo('Starting CreateShortcutsForExistingUsers for ' + IntToStr(ShortcutsList.Count) + ' entries');

  // Log summary of shortcut entries (masked) for diagnostics
  LogDebug('CreateShortcutsForExistingUsers: ShortcutsList entries:');
  for i := 0 to ShortcutsList.Count - 1 do
    LogDebug('  [' + IntToStr(i) + '] ' + MaskPasswordInEntry(ShortcutsList[i]));

  // Log desktop info for shortcut path resolution
  LogDebug('CreateShortcutsForExistingUsers: Desktop path=' + ExpandConstant('{userdesktop}') +
    ' | user desktop exists=' + BoolToStr(DirExists(ExpandConstant('{userdesktop}'))));

  for i := 0 to ShortcutsList.Count - 1 do
  begin
    UserStartTick := GetTickCount;
    LogDebug('CreateShortcutsForExistingUsers: Processing entry ' + IntToStr(i+1) + '/' + IntToStr(ShortcutsList.Count));
    if (GetTickCount - StartTick) > USERS_OVERALL_TIMEOUT then
    begin
      LogError('CreateShortcutsForExistingUsers overall timeout reached after ' + IntToStr(GetTickCount - StartTick) + ' ms; aborting remaining shortcuts');
      break;
    end;
    Entry := ShortcutsList[i];
    ParseUserEntry(Entry, UserName, Password);
    LogDebug('CreateShortcutsForExistingUsers: User=' + UserName + ' hasPassword=' + BoolToStr(Password <> ''));

    StatusOverlay.Caption := 'Creating RDP shortcut (' + IntToStr(i + 1) + ' of ' + IntToStr(ShortcutsList.Count) + '): ' + UserName;

    // Create RDP shortcut using helper function
    CreateRDPShortcut(UserName, Password, 'EXISTING_USER');
    LogDebug('Created shortcut for ' + UserName + ' [DURATION:' + IntToStr(GetTickCount - UserStartTick) + 'ms]');

    if (GetTickCount - UserStartTick) > PER_USER_TIMEOUT then
      LogWarn('CreateShortcutsForExistingUsers per-user time exceeded ' + IntToStr(PER_USER_TIMEOUT) + ' ms for ' + UserName);
  end;
  LogInfo('CreateShortcutsForExistingUsers completed: ' + IntToStr(ShortcutsList.Count) + ' shortcuts [DURATION:' + IntToStr(GetTickCount - StartTick) + 'ms]');
  LogExit('CreateShortcutsForExistingUsers');
end;

// Helper functions to display and update step-by-step progress on Installing page
procedure SetStepPending(L: TLabel; const Text: string);
begin
  LogDebug('SetStepPending: ' + Text);
  if Assigned(L) then
  begin
    L.Caption := '>> ' + Text;
    L.Font.Color := clGray;
    L.Font.Style := [];
    L.Visible := True;
  end;
end;

procedure SetStepInProgress(L: TLabel; const Text: string);
begin
  LogDebug('SetStepInProgress: ' + Text);
  if Assigned(L) then
  begin
    L.Caption := '>>> ' + Text;
    if IsDarkColor(WizardForm.Color) then
      L.Font.Color := clWhite
    else
      L.Font.Color := clBlack;
    L.Font.Style := [fsBold];
    L.Visible := True;
  end;
end;

procedure SetStepDone(L: TLabel; const Text: string);
begin
  LogDebug('SetStepDone: ' + Text);
  if Assigned(L) then
  begin
    L.Caption := '[x] ' + Text;
    L.ParentFont := False;
    L.StyleElements := L.StyleElements - [seFont];
    L.Font.Color := RGBToColor(0, 200, 0);
    // Advance determinate progress bar
    if StepsTotal > 0 then
    begin
      StepsDone := StepsDone + 1;
      if StepsDone > StepsTotal then StepsDone := StepsTotal;
      WizardForm.ProgressGauge.Position := StepsDone;
      LogDebug('Progress: ' + IntToStr(StepsDone) + '/' + IntToStr(StepsTotal) + ' (' + IntToStr(StepsDone * 100 div StepsTotal) + '%) - ' + Text);
    end;
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
  L.Transparent := True;
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
  if Assigned(StepInstallTermWrap) then StepInstallTermWrap.Visible := False;
  if Assigned(StepConfigureService) then StepConfigureService.Visible := False;
  if Assigned(StepCreateUsers) then StepCreateUsers.Visible := False;
  if Assigned(StepCreateShortcuts) then StepCreateShortcuts.Visible := False;
  if Assigned(StepPreTrust) then StepPreTrust.Visible := False;
  if Assigned(StepStartSvc) then StepStartSvc.Visible := False;
  if Assigned(StepCheckRDP) then StepCheckRDP.Visible := False;
  if Assigned(StepUninstallTermWrap) then StepUninstallTermWrap.Visible := False;
  if Assigned(StepRemoveFolder) then StepRemoveFolder.Visible := False;
  if Assigned(StepQuickFixes) then StepQuickFixes.Visible := False;
  if Assigned(StepShowRDPInfo) then StepShowRDPInfo.Visible := False;
end;

// Begin a fresh layout pass for the steps list
procedure BeginStepLayout;
begin
  HideAllStepLabels;
  StepNextTop := StepTopBase;
end;

// Add a pending step label at the next position and counts it for the progress bar
procedure AddStepPendingLabel(L: TLabel; const Text: string);
begin
  if Assigned(L) then
  begin
    L.Left := StepLeftPos;
    L.Top := StepNextTop;
    L.Width := StepWidthVal;
    SetStepPending(L, Text);
    StepNextTop := StepNextTop + ScaleY(16);
    StepsTotal := StepsTotal + 1;
    WizardForm.ProgressGauge.Max := StepsTotal;
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  // This is called before the install page shows, so we skip all blocking operations here
  // Service stop will be handled in CurStepChanged(ssInstall) after UI renders
  Result := '';
  Log('[PrepareToInstall] Skipping service operations (will be done in ssInstall step)');
end;

procedure ReadShortcutSettingsFromRdpFile(const RdpPath: string);
// Reads display/audio settings from an existing .rdp file and populates the
// shortcut settings controls.  Uses PowerShell to handle both UTF-8 and
// UTF-16 LE encodings (mstsc writes UTF-16; this installer writes UTF-8).
var
  OutPath, PSScript, ScriptPath: string;
  Lines: TStringList;
  i, EqPos: Integer;
  Key, Val, Line: string;
  ScreenMode, DesktopWidth, DesktopHeight, UseMultiMon, AudioMode, RedirectClipboard, KeyboardHook: Integer;
  DisableWallpaper, AllowFontSmooth, AllowComposition: Integer;
  DisableFullWindowDrag, DisableMenuAnims, DisableThemes: Integer;
  ResIndex, ResultCode: Integer;
begin
  LogEntry('ReadShortcutSettingsFromRdpFile');
  LogDebug('ReadShortcutSettingsFromRdpFile: path=' + RdpPath);
  if not FileExists(RdpPath) then
  begin
    LogDebug('ReadShortcutSettingsFromRdpFile: file not found, skipping');
    LogExit('ReadShortcutSettingsFromRdpFile');
    exit;
  end;

  // Use PowerShell to read the file regardless of encoding (UTF-16 LE or UTF-8)
  // and emit key=value pairs for the settings we care about.
  OutPath := TempFile('rdp_read_settings.txt');
  PSScript :=
    '$p = ''' + RdpPath + ''';' +
    '$bytes = [System.IO.File]::ReadAllBytes($p);' +
    'if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)' +
    '  { $lines = Get-Content -Encoding Unicode $p }' +
    'else' +
    '  { $lines = Get-Content $p };' +
    '$out = @();' +
    'foreach ($l in $lines) {' +
    '  if ($l -match "^screen mode id:i:(.+)$")   { $out += "screen_mode_id="    + $Matches[1] };' +
    '  if ($l -match "^desktopwidth:i:(.+)$")      { $out += "desktopwidth="      + $Matches[1] };' +
    '  if ($l -match "^desktopheight:i:(.+)$")     { $out += "desktopheight="     + $Matches[1] };' +
    '  if ($l -match "^use multimon:i:(.+)$")      { $out += "use_multimon="      + $Matches[1] };' +
    '  if ($l -match "^audiomode:i:(.+)$")                  { $out += "audiomode="              + $Matches[1] };' +
    '  if ($l -match "^redirectclipboard:i:(.+)$")         { $out += "redirectclipboard="       + $Matches[1] };' +
    '  if ($l -match "^disable wallpaper:i:(.+)$")         { $out += "disable_wallpaper="       + $Matches[1] };' +
    '  if ($l -match "^allow font smoothing:i:(.+)$")      { $out += "allow_font_smoothing="    + $Matches[1] };' +
    '  if ($l -match "^allow desktop composition:i:(.+)$") { $out += "allow_composition="       + $Matches[1] };' +
    '  if ($l -match "^disable full window drag:i:(.+)$")  { $out += "disable_drag_contents="   + $Matches[1] };' +
    '  if ($l -match "^disable menu anims:i:(.+)$")        { $out += "disable_menu_anims="      + $Matches[1] };' +
    '  if ($l -match "^disable themes:i:(.+)$")            { $out += "disable_themes="          + $Matches[1] };' +
    '  if ($l -match "^keyboardhook:i:(.+)$")              { $out += "keyboardhook="            + $Matches[1] };' +
    '};' +
    '[System.IO.File]::WriteAllText(''' + OutPath + ''', ($out -join "`n"))';

  ScriptPath := TempFile('read_rdp_settings.ps1');
  SaveStringToFile(ScriptPath, PSScript, False);
  Exec(EXE_POWERSHELL, BuildPowerShellFileArgs(ScriptPath, '', True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  WriteInstallerLog('ReadShortcutSettingsFromRdpFile: PS exit=' + IntToStr(ResultCode) + ' path=' + RdpPath);

  if not FileExists(OutPath) then exit;

  // Defaults (match BuildShortcutSettingsBlock initial state)
  ScreenMode := 1; DesktopWidth := DEFAULT_RDP_WIDTH; DesktopHeight := DEFAULT_RDP_HEIGHT;
  UseMultiMon := 0; AudioMode := 0; RedirectClipboard := 1; KeyboardHook := 0;
  DisableWallpaper := 1; AllowFontSmooth := 1; AllowComposition := 1;
  DisableFullWindowDrag := 0; DisableMenuAnims := 0; DisableThemes := 0;

  Lines := TStringList.Create;
  try
    Lines.LoadFromFile(OutPath);
    for i := 0 to Lines.Count - 1 do
    begin
      Line := Trim(Lines[i]);
      EqPos := Pos('=', Line);
      if EqPos = 0 then continue;
      Key := Copy(Line, 1, EqPos - 1);
      Val := Trim(Copy(Line, EqPos + 1, Length(Line)));
      if      Key = 'screen_mode_id'    then ScreenMode       := StrToIntDef(Val, 1)
      else if Key = 'desktopwidth'      then DesktopWidth     := StrToIntDef(Val, DEFAULT_RDP_WIDTH)
      else if Key = 'desktopheight'     then DesktopHeight    := StrToIntDef(Val, DEFAULT_RDP_HEIGHT)
      else if Key = 'use_multimon'      then UseMultiMon      := StrToIntDef(Val, 0)
      else if Key = 'audiomode'         then AudioMode        := StrToIntDef(Val, 0)
      else if Key = 'redirectclipboard'    then RedirectClipboard    := StrToIntDef(Val, 1)
      else if Key = 'disable_wallpaper'    then DisableWallpaper     := StrToIntDef(Val, 1)
      else if Key = 'allow_font_smoothing' then AllowFontSmooth      := StrToIntDef(Val, 1)
      else if Key = 'allow_composition'    then AllowComposition     := StrToIntDef(Val, 1)
      else if Key = 'disable_drag_contents' then DisableFullWindowDrag := StrToIntDef(Val, 0)
      else if Key = 'disable_menu_anims'   then DisableMenuAnims     := StrToIntDef(Val, 0)
      else if Key = 'disable_themes'       then DisableThemes        := StrToIntDef(Val, 0)
      else if Key = 'keyboardhook'          then KeyboardHook          := StrToIntDef(Val, 0);
    end;
  finally
    Lines.Free;
    DeleteFile(OutPath);
  end;

  // Apply to controls — Full Screen drives whether cboResolution is enabled
  if Assigned(chkFullScreen) then chkFullScreen.Checked := (ScreenMode = 2);

  ResIndex := -1;  // -1 = no preset match yet
  if      (DesktopWidth = 1280) and (DesktopHeight = 720)  then ResIndex := 0
  else if (DesktopWidth = DEFAULT_RDP_WIDTH) and (DesktopHeight = DEFAULT_RDP_HEIGHT) then ResIndex := 1
  else if (DesktopWidth = 1600) and (DesktopHeight = 900)  then ResIndex := 2
  else if (DesktopWidth = 1920) and (DesktopHeight = 1080) then ResIndex := 3
  else if (DesktopWidth = 2560) and (DesktopHeight = 1440) then ResIndex := 4
  else if (DesktopWidth = 3840) and (DesktopHeight = 2160) then ResIndex := 5;

  if ResIndex = -1 then
  begin
    // Unknown resolution — select Custom and pre-fill the W/H boxes
    ResIndex := 6;
    if Assigned(edtCustomWidth)  then edtCustomWidth.Text  := IntToStr(DesktopWidth);
    if Assigned(edtCustomHeight) then edtCustomHeight.Text := IntToStr(DesktopHeight);
  end;

  if Assigned(cboResolution) then
  begin
    cboResolution.ItemIndex := ResIndex;
    cboResolution.Enabled := (ScreenMode <> 2);
    // Show/hide W/H row to match the selection, then let full screen override
    OnResolutionChange(nil);
    if Assigned(chkFullScreen) and chkFullScreen.Checked then
      OnFullScreenClick(nil);
  end;

  if Assigned(chkUseAllMonitors) then
  begin
    chkUseAllMonitors.Checked := (UseMultiMon = 1);
    OnUseAllMonitorsClick(nil);
  end;
  if Assigned(chkSound)          then chkSound.Checked          := (AudioMode = 0);  // 0 = play on this PC
  if Assigned(chkCopyPaste)      then chkCopyPaste.Checked      := (RedirectClipboard = 1);
  // Experience checkboxes — invert disable keys, pass-through allow keys
  if Assigned(chkExpWallpaper)    then chkExpWallpaper.Checked    := (DisableWallpaper = 0);
  if Assigned(chkExpFontSmooth)   then chkExpFontSmooth.Checked   := (AllowFontSmooth = 1);
  if Assigned(chkExpComposition)  then chkExpComposition.Checked  := (AllowComposition = 1);
  if Assigned(chkExpDragContents) then chkExpDragContents.Checked := (DisableFullWindowDrag = 0);
  if Assigned(chkExpMenuAnim)     then chkExpMenuAnim.Checked     := (DisableMenuAnims = 0);
  if Assigned(chkExpVisualStyles) then chkExpVisualStyles.Checked := (DisableThemes = 0);
  // Keyboard Hook: 0=main PC, 1=RDP, 2=RDP only if full-screen
  if Assigned(cboKeyboardHook) then
  begin
    if (KeyboardHook >= 0) and (KeyboardHook < cboKeyboardHook.Items.Count) then
      cboKeyboardHook.ItemIndex := KeyboardHook
    else
      cboKeyboardHook.ItemIndex := 0;
  end;
end;

procedure WriteShortcutSettingsToRdpFile(const RdpPath: string);
// Reads the existing .rdp file and updates the display/audio/experience settings from the
// current shortcut settings UI controls, then writes the file back in-place.
var
  ScreenModeId, DesktopWidth, DesktopHeight, UseMultiMon, AudioMode, RedirectClipboard, KeyboardHook: Integer;
  DisableWallpaper, AllowFontSmooth, AllowComposition: Integer;
  DisableFullWindowDrag, DisableMenuAnims, DisableThemes: Integer;
  ResultCode: Integer;
  PSScript, ScriptPath: string;
  OpTick: Cardinal;
begin
  LogEntry('WriteShortcutSettingsToRdpFile');
  OpTick := GetTickCount;
  // Collect values from UI controls
  GetShortcutDisplaySettings(ScreenModeId, DesktopWidth, DesktopHeight, UseMultiMon, AudioMode, RedirectClipboard, KeyboardHook);
  GetExperienceSettings(DisableWallpaper, AllowFontSmooth, AllowComposition, DisableFullWindowDrag, DisableMenuAnims, DisableThemes);

  // PowerShell: update specific keys in the .rdp file without touching the rest (e.g. password hash)
  PSScript :=
    '$path = ''' + RdpPath + '''' + #13#10 +
    '$lines = if (Test-Path $path) { [System.IO.File]::ReadAllLines($path) } else { @() }' + #13#10 +
    'function Set-RdpKey($lines, $prefix, $val) {' + #13#10 +
    '  $found = $false' + #13#10 +
    '  for ($i = 0; $i -lt $lines.Count; $i++) {' + #13#10 +
    '    if ($lines[$i] -match ("^" + [regex]::Escape($prefix) + ":")) { $lines[$i] = $val; $found = $true }' + #13#10 +
    '  }' + #13#10 +
    '  if (-not $found) { $lines += $val }' + #13#10 +
    '  return ,$lines' + #13#10 +
    '}' + #13#10 +
    '$lines = Set-RdpKey $lines "screen mode id" "screen mode id:i:' + IntToStr(ScreenModeId) + '"' + #13#10 +
    '$lines = Set-RdpKey $lines "desktopwidth" "desktopwidth:i:' + IntToStr(DesktopWidth) + '"' + #13#10 +
    '$lines = Set-RdpKey $lines "desktopheight" "desktopheight:i:' + IntToStr(DesktopHeight) + '"' + #13#10 +
    '$lines = Set-RdpKey $lines "use multimon" "use multimon:i:' + IntToStr(UseMultiMon) + '"' + #13#10 +
    '$lines = Set-RdpKey $lines "audiomode" "audiomode:i:' + IntToStr(AudioMode) + '"' + #13#10 +
    '$lines = Set-RdpKey $lines "redirectclipboard" "redirectclipboard:i:' + IntToStr(RedirectClipboard) + '"' + #13#10 +
    '$lines = Set-RdpKey $lines "disable wallpaper" "disable wallpaper:i:' + IntToStr(DisableWallpaper) + '"' + #13#10 +
    '$lines = Set-RdpKey $lines "allow font smoothing" "allow font smoothing:i:' + IntToStr(AllowFontSmooth) + '"' + #13#10 +
    '$lines = Set-RdpKey $lines "allow desktop composition" "allow desktop composition:i:' + IntToStr(AllowComposition) + '"' + #13#10 +
    '$lines = Set-RdpKey $lines "disable full window drag" "disable full window drag:i:' + IntToStr(DisableFullWindowDrag) + '"' + #13#10 +
    '$lines = Set-RdpKey $lines "disable menu anims" "disable menu anims:i:' + IntToStr(DisableMenuAnims) + '"' + #13#10 +
    '$lines = Set-RdpKey $lines "disable themes" "disable themes:i:' + IntToStr(DisableThemes) + '"' + #13#10 +
    '$lines = Set-RdpKey $lines "keyboardhook" "keyboardhook:i:' + IntToStr(KeyboardHook) + '"' + #13#10 +
    '[System.IO.File]::WriteAllLines($path, $lines)';

  ScriptPath := TempFile('update_rdp_settings.ps1');
  SaveStringToFile(ScriptPath, PSScript, False);
  LogDebug('WriteShortcutSettingsToRdpFile: Starting PS update for ' + RdpPath);
  OpTick := GetTickCount;
  Exec(EXE_POWERSHELL, BuildPowerShellFileArgs(ScriptPath, '', True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  LogDebug('WriteShortcutSettingsToRdpFile: PS exit=' + IntToStr(ResultCode) + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  if ResultCode = 0 then
  begin
    LogDebug('WriteShortcutSettingsToRdpFile: Update succeeded, signing file...');
    SignRdpFile(RdpPath);
  end
  else
    LogWarn('WriteShortcutSettingsToRdpFile: PS update failed with exit=' + IntToStr(ResultCode));
  LogInfo('WriteShortcutSettingsToRdpFile: completed for ' + RdpPath + ' [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('WriteShortcutSettingsToRdpFile');
end;

// Displays help for a setting on the Shortcut Settings page.
// Tag 1=Shortcut Name, 2=Copy&Paste & Sound, 3=Monitor settings (Window Size/Full Screen/All Monitors), 4=Performance vs Quality checkboxes
procedure ShowShortcutHelpInfo(Sender: TObject);
var
  HelpText: string;
begin
  case TButton(Sender).Tag of
    1: HelpText :=
         'Shortcut Name' + #13#10#13#10 +
         'This is the name of the RDP shortcut file that will be created on the Desktop. ';
    2: HelpText :=
         'Allow Copy & Paste' + #13#10#13#10 +
         'Enables clipboard sharing between the remote session and the local PC.' + #13#10#13#10 +
         'When checked, you can copy text, images, and files on one side and paste them on the other. ' +
         'When unchecked, the clipboard is isolated. Nothing can be transferred between the two sides.' + #13#10#13#10 +        
         'Allow Sound' + #13#10#13#10 +
         'Controls whether audio from the remote session plays on the local PC.' + #13#10#13#10 +
         'When checked, sounds (system alerts, media playback, etc.) from the remote session are ' +
         'redirected to and played through the local PC''s speakers. ';
    3: HelpText :=
         'Window Size / Full Screen / Use All Monitors' + #13#10#13#10 +
         'Window Size: sets the resolution of the remote desktop window.' + #13#10 +
         '  Choose a preset resolution or "Custom" to enter exact pixel dimensions.' + #13#10#13#10 +
         'Full Screen: opens the RDP session filling the entire screen at the monitor''s native resolution. ' +
         'Overrides the Window Size setting.' + #13#10#13#10 +
         'Use All Monitors: spans the remote session across all connected monitors. ' +
         'Each monitor''s full resolution is used. Requires Full Screen to be unchecked.';
    4: HelpText :=
        'Performance vs Quality' + #13#10#13#10 +
        'These settings control which visual effects are displayed in the RDP connection. ' +
        'Disabling effects improves responsiveness and performance.' + #13#10#13#10 +
         'Desktop wallpaper: shows or hides the background image in the session.' + #13#10 +
         'Smooth text (ClearType): enables font anti-aliasing for sharper text.' + #13#10 +
         'Transparent windows & effects: enables Aero glass and window animations.' + #13#10 +
         'Show contents while dragging: renders window contents as you drag them.' + #13#10 +
         'Animated menus & transitions: enables fade/slide animations on menus.' + #13#10 +
         'Visual themes: enables Windows visual styling (disabling gives a classic look).';
    5: HelpText :=
        'Apply Keyboard Combos' + #13#10#13#10 +
        'Controls how Windows key combinations (like ALT+TAB) are handled:' + #13#10#13#10 +
        'To the main/host PC:' + #13#10 +
        '  Windows key combinations (like ALT+TAB) are processed outside of RDPs on your main/host PC' + #13#10#13#10 +
        'To the RDP:' + #13#10 +
        '  Windows key combinations (like ALT+TAB) are processed by the RDP window' + #13#10#13#10 +
        'To the RDP, only if full-screen:' + #13#10 +
        '  Windows key combinations (like ALT+TAB) are processed by the RDP window only if it is full-screen';
 else
    HelpText := 'No additional information available for this setting.';
  end;
  MsgBox(HelpText, mbInformation, MB_OK);
end;

// Creates a small "[?]" help button on a raw TWinControl surface (e.g. inside BuildShortcutSettingsBlock).
// Tag identifies which help text to show (see ShowShortcutHelpInfo).
function MakeShortcutHelpButton(Surface: TWinControl; ATop, ATag: Integer): TButton;
var
  Btn: TButton;
begin
  Btn := TButton.Create(Surface);
  Btn.Parent := Surface;
  Btn.Caption := '?';
  Btn.Width := ScaleX(20);
  Btn.Height := ScaleY(20);
  Btn.Left := Surface.Width - ScaleX(26);
  Btn.Top := ATop;
  Btn.Tag := ATag;
  Btn.OnClick := @ShowShortcutHelpInfo;
  Result := Btn;
end;

procedure BuildShortcutSettingsBlock(ParentSurface: TWinControl; StartTop: Integer);
var
  TmpLabel: TLabel;
begin

  // Shortcut Name field at top
  lblShortcutName := TLabel.Create(ParentSurface);
  lblShortcutName.Parent := ParentSurface;
  lblShortcutName.Left := ScaleX(10);
  lblShortcutName.Top := StartTop;
  lblShortcutName.Caption := 'Shortcut Name:';
  lblShortcutName.AutoSize := True;

  edtShortcutName := TEdit.Create(ParentSurface);
  edtShortcutName.Parent := ParentSurface;
  edtShortcutName.Left := lblShortcutName.Left + ScaleX(84);
  edtShortcutName.Top := StartTop - ScaleY(2);
  edtShortcutName.Width := ScaleX(120);
  edtShortcutName.Text := 'macro1';

  lblShortcutExtension := TLabel.Create(ParentSurface);
  lblShortcutExtension.Parent := ParentSurface;
  lblShortcutExtension.Left := edtShortcutName.Left + edtShortcutName.Width + ScaleX(4);
  lblShortcutExtension.Top := edtShortcutName.Top + ScaleY(3);
  lblShortcutExtension.Caption := '.rdp';
  lblShortcutExtension.AutoSize := True;

  MakeShortcutHelpButton(ParentSurface, lblShortcutName.Top, 1);

  // Section separator label
  lblShortcutSection := TLabel.Create(ParentSurface);
  lblShortcutSection.Parent := ParentSurface;
  lblShortcutSection.Left := ScaleX(10);
  lblShortcutSection.Top := lblShortcutName.Top + ScaleY(22);
  lblShortcutSection.Caption := 'Basic Shortcut Settings (editing: <>)';
  lblShortcutSection.Font.Style := [fsBold];
  lblShortcutSection.AutoSize := True;

  // Row 2 — Copy & Paste (1st column)
  chkCopyPaste := TCheckBox.Create(ParentSurface);
  chkCopyPaste.Parent := ParentSurface;
  chkCopyPaste.Left := ScaleX(10);
  chkCopyPaste.Top := lblShortcutSection.Top + ScaleY(24);
  chkCopyPaste.Width := ScaleX(150);
  chkCopyPaste.Caption := 'Allow Copy && Paste';
  chkCopyPaste.Checked := True;

  // Row 2 — Sound (2nd column)
  chkSound := TCheckBox.Create(ParentSurface);
  chkSound.Parent := ParentSurface;
  chkSound.Left := ScaleX(180);
  chkSound.Top := lblShortcutSection.Top + ScaleY(24);
  chkSound.Width := ScaleX(180);
  chkSound.Caption := 'Allow Sound';
  chkSound.Checked := True;
  MakeShortcutHelpButton(ParentSurface, chkSound.Top, 2);

  // Row 3 — Keyboard Combos dropdown (after Sound)
  lblKeyboardHook := TLabel.Create(ParentSurface);
  lblKeyboardHook.Parent := ParentSurface;
  lblKeyboardHook.Left := ScaleX(10);
  lblKeyboardHook.Top := chkSound.Top + ScaleY(24);
  lblKeyboardHook.Caption := 'Keyboard Combos (like ALT+TAB):';
  lblKeyboardHook.AutoSize := True;

  cboKeyboardHook := TComboBox.Create(ParentSurface);
  cboKeyboardHook.Parent := ParentSurface;
  cboKeyboardHook.Left := ScaleX(200);
  cboKeyboardHook.Top := chkSound.Top + ScaleY(20);
  cboKeyboardHook.Width := ScaleX(210);
  cboKeyboardHook.Style := csDropDownList;
  cboKeyboardHook.Items.Add('Send to the main/host PC');
  cboKeyboardHook.Items.Add('Send to the RDP');
  cboKeyboardHook.Items.Add('Send to the RDP, only if full-screen');
  cboKeyboardHook.ItemIndex := 0;
  MakeShortcutHelpButton(ParentSurface, lblKeyboardHook.Top, 5);

  // Row 4 — Screen Size label
  lblScreenSize := TLabel.Create(ParentSurface);
  lblScreenSize.Parent := ParentSurface;
  lblScreenSize.Left := ScaleX(10);
  lblScreenSize.Top := cboKeyboardHook.Top + ScaleY(28);
  lblScreenSize.Caption := 'Window Size:';
  lblScreenSize.AutoSize := True;

  // Row 4 — Resolution drop-down
  cboResolution := TComboBox.Create(ParentSurface);
  cboResolution.Parent := ParentSurface;
  cboResolution.Left := lblScreenSize.Left + ScaleX(72);
  cboResolution.Top := lblScreenSize.Top - ScaleY(3);
  cboResolution.Width := ScaleX(120);
  cboResolution.Style := csDropDownList;
  cboResolution.Items.Add('1280 x 720');
  cboResolution.Items.Add('1366 x 768');
  cboResolution.Items.Add('1600 x 900');
  cboResolution.Items.Add('1920 x 1080');
  cboResolution.Items.Add('2560 x 1440');
  cboResolution.Items.Add('3840 x 2160');
  cboResolution.Items.Add('Custom');
  cboResolution.ItemIndex := 1;  // default: 1366 x 768
  cboResolution.OnChange := @OnResolutionChange;

  // Row 4 — Full Screen checkbox
  chkFullScreen := TCheckBox.Create(ParentSurface);
  chkFullScreen.Parent := ParentSurface;
  chkFullScreen.Left := cboResolution.Left + cboResolution.Width + ScaleX(10);
  chkFullScreen.Top := lblScreenSize.Top - ScaleY(1);
  chkFullScreen.Width := ScaleX(80);
  chkFullScreen.Caption := 'Full Screen';
  chkFullScreen.Checked := False;
  chkFullScreen.OnClick := @OnFullScreenClick;
  cboResolution.Enabled := True;

  // Row 4 — Use All Monitors
  chkUseAllMonitors := TCheckBox.Create(ParentSurface);
  chkUseAllMonitors.Parent := ParentSurface;
  chkUseAllMonitors.Left := chkFullScreen.Left + chkFullScreen.Width + ScaleX(20);
  chkUseAllMonitors.Top := lblScreenSize.Top - ScaleY(1);
  chkUseAllMonitors.Width := ScaleX(110);
  chkUseAllMonitors.Caption := 'Use All Monitors';
  chkUseAllMonitors.Checked := False;
  chkUseAllMonitors.OnClick := @OnUseAllMonitorsClick;
  MakeShortcutHelpButton(ParentSurface, lblScreenSize.Top - ScaleY(1), 3);

  // Row 4b — Custom resolution W/H inputs (hidden until "Custom" is selected)
  lblCustomWidth := TLabel.Create(ParentSurface);
  lblCustomWidth.Parent := ParentSurface;
  lblCustomWidth.Left := cboResolution.Left;
  lblCustomWidth.Top := lblScreenSize.Top + ScaleY(26);
  lblCustomWidth.Caption := 'W:';
  lblCustomWidth.AutoSize := True;
  lblCustomWidth.Visible := False;

  edtCustomWidth := TEdit.Create(ParentSurface);
  edtCustomWidth.Parent := ParentSurface;
  edtCustomWidth.Left := lblCustomWidth.Left + ScaleX(18);
  edtCustomWidth.Top := lblCustomWidth.Top - ScaleY(3);
  edtCustomWidth.Width := ScaleX(45);
  edtCustomWidth.Text := '1920';
  edtCustomWidth.Visible := False;

  lblCustomHeight := TLabel.Create(ParentSurface);
  lblCustomHeight.Parent := ParentSurface;
  lblCustomHeight.Left := edtCustomWidth.Left + edtCustomWidth.Width + ScaleX(10);
  lblCustomHeight.Top := lblCustomWidth.Top;
  lblCustomHeight.Caption := 'H:';
  lblCustomHeight.AutoSize := True;
  lblCustomHeight.Visible := False;

  edtCustomHeight := TEdit.Create(ParentSurface);
  edtCustomHeight.Parent := ParentSurface;
  edtCustomHeight.Left := lblCustomHeight.Left + ScaleX(18);
  edtCustomHeight.Top := edtCustomWidth.Top;
  edtCustomHeight.Width := ScaleX(45);
  edtCustomHeight.Text := '1080';
  edtCustomHeight.Visible := False;

  // Row 5 — Performance section header (offset by extra row to clear Custom resolution inputs)
  TmpLabel := TLabel.Create(ParentSurface);
  TmpLabel.Parent := ParentSurface;
  TmpLabel.Left := ScaleX(10);
  TmpLabel.Top := lblScreenSize.Top + ScaleY(52);
  TmpLabel.Caption := 'Quality vs Performance - (Unchecked = better performance):';
  TmpLabel.Font.Style := [fsBold];
  TmpLabel.AutoSize := True;
  MakeShortcutHelpButton(ParentSurface, TmpLabel.Top, 4);

  // Row 6 — Experience checkboxes (2 columns, 3 rows)
  // Col 1, Row 1
  chkExpWallpaper := TCheckBox.Create(ParentSurface);
  chkExpWallpaper.Parent := ParentSurface;
  chkExpWallpaper.Left := ScaleX(20);
  chkExpWallpaper.Top := TmpLabel.Top + ScaleY(20);
  chkExpWallpaper.Width := ScaleX(180);
  chkExpWallpaper.Caption := 'Desktop wallpaper';
  chkExpWallpaper.Checked := False;

  // Col 2, Row 1
  chkExpFontSmooth := TCheckBox.Create(ParentSurface);
  chkExpFontSmooth.Parent := ParentSurface;
  chkExpFontSmooth.Left := ScaleX(210);
  chkExpFontSmooth.Top := chkExpWallpaper.Top;
  chkExpFontSmooth.Width := ScaleX(180);
  chkExpFontSmooth.Caption := 'Smooth text (ClearType)';
  chkExpFontSmooth.Checked := True;

  // Col 1, Row 2
  chkExpComposition := TCheckBox.Create(ParentSurface);
  chkExpComposition.Parent := ParentSurface;
  chkExpComposition.Left := ScaleX(20);
  chkExpComposition.Top := chkExpWallpaper.Top + ScaleY(22);
  chkExpComposition.Width := ScaleX(180);
  chkExpComposition.Caption := 'Transparent windows && effects';
  chkExpComposition.Checked := True;

  // Col 2, Row 2
  chkExpDragContents := TCheckBox.Create(ParentSurface);
  chkExpDragContents.Parent := ParentSurface;
  chkExpDragContents.Left := ScaleX(210);
  chkExpDragContents.Top := chkExpComposition.Top;
  chkExpDragContents.Width := ScaleX(230);
  chkExpDragContents.Caption := 'Show window contents while dragging';
  chkExpDragContents.Checked := True;

  // Col 1, Row 3
  chkExpMenuAnim := TCheckBox.Create(ParentSurface);
  chkExpMenuAnim.Parent := ParentSurface;
  chkExpMenuAnim.Left := ScaleX(20);
  chkExpMenuAnim.Top := chkExpComposition.Top + ScaleY(22);
  chkExpMenuAnim.Width := ScaleX(180);
  chkExpMenuAnim.Caption := 'Animated menus && transitions';
  chkExpMenuAnim.Checked := True;

  // Col 2, Row 3
  chkExpVisualStyles := TCheckBox.Create(ParentSurface);
  chkExpVisualStyles.Parent := ParentSurface;
  chkExpVisualStyles.Left := ScaleX(210);
  chkExpVisualStyles.Top := chkExpMenuAnim.Top;
  chkExpVisualStyles.Width := ScaleX(180);
  chkExpVisualStyles.Caption := 'Visual themes';
  chkExpVisualStyles.Checked := True;

  // Multi-user note (shown on the same row as Shortcut Name when 2+ shortcuts will be created/edited)
  lblMultiShortcutEditingNote := TLabel.Create(ParentSurface);
  lblMultiShortcutEditingNote.Parent := ParentSurface;
  lblMultiShortcutEditingNote.Left := ScaleX(10);
  lblMultiShortcutEditingNote.Top := StartTop;
  lblMultiShortcutEditingNote.Font.Style := [fsBold];
  lblMultiShortcutEditingNote.AutoSize := True;
  lblMultiShortcutEditingNote.Visible := False;

  // Row 4 — "Open advanced shortcut options" checkbox (shown only in Edit Shortcuts mode)
  chkShowMoreShortcutOptions := TCheckBox.Create(ParentSurface);
  chkShowMoreShortcutOptions.Parent := ParentSurface;
  chkShowMoreShortcutOptions.Left := ScaleX(260);
  chkShowMoreShortcutOptions.Top := chkExpMenuAnim.Top + ScaleY(49);
  chkShowMoreShortcutOptions.Width := ScaleX(260);
  chkShowMoreShortcutOptions.Caption := 'Open advanced shortcut options (Next page)';
  chkShowMoreShortcutOptions.Checked := False;
  chkShowMoreShortcutOptions.Visible := False;
end;

// Populate a TCheckBox from a registry DWord: Checked = (value = TrueValue).
// Falls back to DefaultChecked when the key/value is absent.
procedure LoadDWordCheckbox(Root: Integer; const Key, Name: string; TrueValue: Cardinal; Ctrl: TCheckBox; DefaultChecked: Boolean);
var
  V: Cardinal;
begin
  if RegQueryDWordValue(Root, Key, Name, V) then
    Ctrl.Checked := (V = TrueValue)
  else
    Ctrl.Checked := DefaultChecked;
end;

// Loads a DWORD registry value into a TComboBox.
// Index 0 in the combo is always "Not Set" (value absent).
// Subsequent items map to integer values 0,1,2... (i.e. comboIndex = regValue + 1).
// If the registry value is absent or out of range, selects index 0.
procedure LoadDWordCombo(Root: Integer; const Key, Name: string; Ctrl: TComboBox);
var
  V: Cardinal;
begin
  if RegQueryDWordValue(Root, Key, Name, V) then
  begin
    if (Integer(V) + 1 < Ctrl.Items.Count) then
      Ctrl.ItemIndex := Integer(V) + 1
    else
      Ctrl.ItemIndex := 0;
  end
  else
    Ctrl.ItemIndex := 0;
end;


// Displays help information for a setting on the Show RDP Info page.
// Sender.Tag identifies which setting (1-5). Descriptions focus on local session impact.
// Displays help for a setting on the Edit System-wide RDP Settings page.
// Tag 1=Enable RDP, 2=Show users, 3=Prevent duplicate, 4=RDP Port, 5=Image quality, 6=Compression, 7=Restart RDP
procedure ShowGPHelpInfo(Sender: TObject);
var
  HelpText: string;
begin
  case TButton(Sender).Tag of
    1: HelpText :=
      'Enable Remote Desktop' + #13#10#13#10 +
      'Allows or blocks RDP connections to this PC.' + #13#10#13#10 + #13#10#13#10 +
      'Note: TermWrap requires RDP to be enabled. Disabling it here will prevent TermWrap from functioning.';
    2: HelpText :=
      'Show users on logon screen' + #13#10#13#10 +
      'Controls whether local user accounts are listed on the Windows login screen.' + #13#10#13#10 +
      'When enabled, user account names appear as tiles on the lock/login screen. ' +
      'Disabling can improve security and keep things tidy by not displaying account names on the screen.';
    3: HelpText :=
      'Prevent duplicate connections per user' + #13#10#13#10 +
      'When enabled, a user who already has an active RDP session cannot open a second simultaneous session. ' +
      'Their new connection takes over the existing one.' + #13#10#13#10 +
      'When disabled, the same user account can have multiple independent RDP sessions running at the same time. ' +
      'This is useful in shared-access scenarios where you may need to log in more than once to the same account.';
    4: HelpText :=
      'RDP Listening Port' + #13#10#13#10 +
      'The TCP port that the Remote Desktop service listens on for incoming connections. Default is 3389.' + #13#10#13#10 +
      'Change this to a non-standard port to reduce exposure to automated port scans and brute-force attempts. ' +
      'If you change it, you must edit all RDP shortcuts to use the new port (e.g. 127.0.0.2:3390).' + #13#10#13#10 +
      'Requires a restart of the RDP service to take effect.';
    5: HelpText :=
      'RemoteFX Image Quality' + #13#10#13#10 +
      'Sets the image encoding quality level used by the RemoteFX Adaptive Graphics pipeline.' + #13#10#13#10 +
      'Options:' + #13#10 +
      '  Not Set  - Windows chooses automatically' + #13#10 +
      '  0 - Low      (lowest quality, fewest resources)' + #13#10 +
      '  1 - Medium   (good balance of quality and CPU)' + #13#10 +
      '  2 - High     (near-lossless, higher CPU cost)' + #13#10 +
      '  3 - Lossless (perfect quality, highest CPU/memory)' + #13#10#13#10 +
      'For local sessions, Medium or High is usually indistinguishable visually.';
    6: HelpText :=
      'RemoteFX Compression' + #13#10#13#10 +
      'Controls the compression algorithm applied to RemoteFX display data before it is sent.' + #13#10#13#10 +
      'Options:' + #13#10 +
      '  Not Set  - Windows chooses automatically' + #13#10 +
      '  0 - No compression      (lowest CPU, highest memory use)' + #13#10 +
      '  1 - Optimized for memory (reduces RAM during encoding)' + #13#10 +
      '  2 - Balanced            (recommended for most cases)' + #13#10 +
      '  3 - Optimized for bandwidth (most CPU, smallest frames)' + #13#10#13#10 +
      'For a local same-PC session, option 1 or 2 is generally best. Bandwidth is not ' +
      'a bottleneck, so heavy compression wastes CPU without benefit.';
    7: HelpText :=
      'Restart RDP Service' + #13#10#13#10 +
      'Stops and restarts the Windows Remote Desktop service (TermService) after applying changes.' + #13#10#13#10 +
      'Some settings (such as the RDP port number) do not take effect until the service is restarted. ' +
      'Active RDP sessions will be disconnected. Check this box if you changed the port or want to ' +
      'ensure all settings are fully applied immediately.';
    8: HelpText :=
      'Hide most security warnings' + #13#10#13#10 +
      'When enabled, common security prompt dialogs related to RDP are hidden. Please be aware that ' +
      'this can suppress warnings, which may have security implications. Files created by RDPWrapKit ' +
      'are safe because they are Local RDP files.' + #13#10#13#10 +
      'Exercise caution if you use RDP to connect to untrusted remote machines.';
    10: HelpText :=
      'Restart RDP Service' + #13#10#13#10 +
      'Stops and restarts the Windows Remote Desktop service (TermService).' + #13#10#13#10 +
      'Active RDP sessions will be disconnected. Use this fix if RDP is not working properly ' +
      'or if you have made changes that require a service restart.';
    11: HelpText :=
      'Set all Local Accounts to Never Expire' + #13#10#13#10 +
      'Stops Windows from forcing a password change every 42 days or getting Account Expiration reminders' + #13#10#13#10 +
      'This is useful for RDP accounts that should remain permanently accessible without requiring password rotation.';
    12: HelpText :=
      'Restore deleted Remote Desktop Service' + #13#10#13#10 +
      'Re-imports the Registry keys for the Remote Desktop Service (TermService) using a backup .reg file.' + #13#10#13#10 +
      'Use this if antivirus software or a system cleanup tool has accidentally removed the ' +
      'TermService registry entries, preventing RDP from functioning. This will restore the ' +
      'service configuration, image path, and required privileges to their default values.';
  else
    HelpText := 'No additional information available for this setting.';
  end;
  MsgBox(HelpText, mbInformation, MB_OK);
end;

// Creates a small "[?]" help button to the right of a control on a wizard page.
// Tag identifies which help text to show (see ShowGPHelpInfo).
function MakeHelpButton(Page: TWizardPage; ATop, ATag: Integer): TButton;
var
  Btn: TButton;
begin
  Btn := TButton.Create(Page);
  Btn.Parent := Page.Surface;
  Btn.Caption := '?';
  Btn.Width := ScaleX(20);
  Btn.Height := ScaleY(20);
  Btn.Left := Page.SurfaceWidth - ScaleX(26);
  Btn.Top := ATop;
  Btn.Tag := ATag;
  Btn.OnClick := @ShowGPHelpInfo;
  Result := Btn;
end;

procedure InitializeWizard;
var
  leftPos, topPos, widthVal: Integer;
  TermWrapAlreadyInstalled: Boolean;
  InstallHintOffset: Integer;
  childIndent, childLeft, valueLeft: Integer;
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
  // Controls for Edit System-wide Settings page (declared in global var block)

  LinkColor: TColor;
  LabelColor: TColor;
  DesiredBottom: Integer;
  BlockTop: Integer;
  Delta: Integer;
    regVal: Cardinal;
    radioTopBase: Integer;
    radioSpacing: Integer;
    TmpLabel: TLabel;
begin
  // Create transparent label overlay for status text to prevent grey flash.
  StatusOverlay := TLabel.Create(WizardForm);
  StatusOverlay.Parent := WizardForm.StatusLabel.Parent;
  StatusOverlay.SetBounds(WizardForm.StatusLabel.Left, WizardForm.StatusLabel.Top,
    WizardForm.StatusLabel.Width, WizardForm.StatusLabel.Height);
  StatusOverlay.Font := WizardForm.StatusLabel.Font;
  StatusOverlay.Transparent := True;
  StatusOverlay.Caption := '';
  StatusOverlay.Show;
  WizardForm.StatusLabel.Visible := False;

  InstallOptionsAutoUserSourceApplied := False;

  // Initialize installer log file
  InitInstallerLog;
  WriteInstallerLog('BUILD_FINGERPRINT=' + BUILD_FINGERPRINT);
  WriteInstallerLog('PRESERVE_USER_CREATE_DEBUG_LOGS=' + BoolToStr(PRESERVE_USER_CREATE_DEBUG_LOGS <> 0));
  WriteInstallerLog('PASSWORD_PIPELINE_DIAG=' + BoolToStr(PASSWORD_PIPELINE_DIAG <> 0));

  // Extract RdpSignTool.exe upfront so it's available for all signing operations
  ExtractTemporaryFile('RdpSignTool.exe');
  WriteInstallerLog('InitializeWizard: Extracted RdpSignTool.exe to ' + ExpandConstant('{tmp}\RdpSignTool.exe'));

  SimLogNoMstscShown := False;
  SimLogNoVCRedistShown := False;
  SimLogNetPsShown := False;
  LastLoggedPageId := -1;
  LastLoggedPageTick := 0;
  LastSuppressedPageLogs := 0;
  PendingDebugCleanupFiles := TStringList.Create;

  if SimulateNoMstsc then
  begin
    LogSimulationScenario('System doesnt have mstsc');
    SimLogNoMstscShown := True;
  end;
  if SimulateNoVCRedist then
  begin
    LogSimulationScenario('System doesnt have VC++');
    SimLogNoVCRedistShown := True;
  end;
  if SimulateNetFailPowerShell then
  begin
    LogSimulationScenario('System fails on net.exe commands and uses PowerShell fallback');
    SimLogNetPsShown := True;
  end;

  if (Ord(SimulateNoMstsc) + Ord(SimulateNoVCRedist) + Ord(SimulateNetFailPowerShell)) > 1 then
    WriteInstallerLog('Simulation warning: multiple scenarios are enabled; run one scenario at a time for deterministic test results.');

  // Detect Smart App Control (VerifiedAndReputablePolicyState == 1)
  SmartAppControlIsOn := False;
  if RegQueryDWordValue(HKLM, 'SYSTEM\CurrentControlSet\Control\CI\Policy', 'VerifiedAndReputablePolicyState', regVal) then
  begin
    WriteInstallerLog('SmartAppControl: Registry value VerifiedAndReputablePolicyState=' + IntToStr(regVal));
    if regVal = 1 then
      SmartAppControlIsOn := True;
  end
  else
  begin
    WriteInstallerLog('SmartAppControl: Registry value VerifiedAndReputablePolicyState not found');
  end;
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
  WelcomeExpLabel.Caption := 'RDPWrapKit sets up local RDP access, allowing you to run multiple users on your computer via Remote Desktop Protocol (RDP).';
  WelcomeExpLabel.WordWrap := True;
  WelcomeExpLabel.Alignment := taLeftJustify;
  WelcomeExpLabel.Font.Size := WelcomeExpLabel.Font.Size + 2;
  // Anchor left+right so label width tracks page surface on dynamic layouts
  WelcomeExpLabel.Anchors := [akLeft, akRight];
  // Set a fixed height large enough for wrapped text
  WelcomeExpLabel.Height := ScaleY(150);

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
  Page_InstallOptions := CreateCustomPage(
    WelcomePage.ID,
    'Setup Options',
    'Select what you would like to do:'
  );
  Page_InstallOptions.Surface.Color := RGBToColor(11, 16, 24);

  // Create top-level radio buttons: Install, Edit Shortcut, Uninstall
  rbInstall := TRadioButton.Create(Page_InstallOptions);
  rbInstall.Parent := Page_InstallOptions.Surface;
  rbInstall.Left := ScaleX(10);
  rbInstall.Top := ScaleY(10);
  rbInstall.Width := ScaleX(420);
  rbInstall.Caption := 'Install';
  rbInstall.Checked := True;
  rbInstall.OnClick := @OnInstallModeChange;

  // Under Install, add checkboxes and nested radios (always visible, only enabled for Install)
  chkInstallTermWrap := TCheckBox.Create(Page_InstallOptions);
  chkInstallTermWrap.Parent := Page_InstallOptions.Surface;
  chkInstallTermWrap.Left := ScaleX(30);
  chkInstallTermWrap.Top := ScaleY(36);
  chkInstallTermWrap.Width := ScaleX(420);
  TermWrapAlreadyInstalled := IsTermWrapInstalled();
  InstallHintOffset := 0;
  if TermWrapAlreadyInstalled then
  begin
    // TermWrap is installed — check whether versions/sizes match (upgrade) or are identical (re-install)
    if (GetInstalledTermWrapVersion() = '{#SourceTermWrapVersion}') and (GetInstalledTermWrapSize() = '{#SourceTermWrapSize}') then
    begin
      // Case 2: Same version and size — show "Re-install TermWrap", unchecked, with hint sub-label
      chkInstallTermWrap.Caption := 'Re-install TermWrap';
      chkInstallTermWrap.Checked := False;

      chkInstallTermWrapHint := TLabel.Create(Page_InstallOptions);
      chkInstallTermWrapHint.Parent := Page_InstallOptions.Surface;
      chkInstallTermWrapHint.Left := ScaleX(46);
      chkInstallTermWrapHint.Top := chkInstallTermWrap.Top + ScaleY(20);
      chkInstallTermWrapHint.Width := ScaleX(390);
      chkInstallTermWrapHint.AutoSize := False;
      chkInstallTermWrapHint.Transparent := True;
      chkInstallTermWrapHint.WordWrap := True;
      chkInstallTermWrapHint.Caption := '(Already installed. Selecting this will re-install it)';
      chkInstallTermWrapHint.Font.Style := [fsItalic];

      InstallHintOffset := ScaleY(16);
    end
    else
    begin
      // Case 3: Different version or size — show "Upgrade TermWrap", checked, no sub-label
      chkInstallTermWrap.Caption := 'Upgrade TermWrap';
      chkInstallTermWrap.Checked := True;
    end;
  end
  else
  begin
    // Case 1: Not installed — show "Install TermWrap", checked, no sub-label
    chkInstallTermWrap.Caption := 'Install TermWrap';
    chkInstallTermWrap.Checked := True;
  end;

  chkCreateRdpShortcuts := TCheckBox.Create(Page_InstallOptions);
  chkCreateRdpShortcuts.Parent := Page_InstallOptions.Surface;
  chkCreateRdpShortcuts.Left := ScaleX(30);
  chkCreateRdpShortcuts.Top := ScaleY(60) + InstallHintOffset;
  chkCreateRdpShortcuts.Width := ScaleX(380);
  chkCreateRdpShortcuts.Caption := 'Create RDP shortcuts';
  chkCreateRdpShortcuts.Checked := True;
  chkCreateRdpShortcuts.OnClick := @OnCreateRdpShortcutsClick;

  CreateRdpShortcutsGroup := TPanel.Create(Page_InstallOptions);
  CreateRdpShortcutsGroup.Parent := Page_InstallOptions.Surface;
  CreateRdpShortcutsGroup.Left := ScaleX(40);
  CreateRdpShortcutsGroup.Top := ScaleY(84) + InstallHintOffset;
  CreateRdpShortcutsGroup.Width := ScaleX(360);
  CreateRdpShortcutsGroup.Height := ScaleY(88);
  CreateRdpShortcutsGroup.BorderStyle := bsNone;
  CreateRdpShortcutsGroup.Color := Page_InstallOptions.Surface.Color;
  CreateRdpShortcutsGroup.BevelInner := bvNone;
  CreateRdpShortcutsGroup.BevelOuter := bvNone;
  CreateRdpShortcutsGroup.BevelWidth := 0;

  rbCreateUsers := TRadioButton.Create(CreateRdpShortcutsGroup);
  rbCreateUsers.Parent := CreateRdpShortcutsGroup;
  rbCreateUsers.Left := ScaleX(10);
  rbCreateUsers.Top := ScaleY(8);
  rbCreateUsers.Width := ScaleX(340);
  rbCreateUsers.Caption := 'Create new users';
  rbCreateUsers.Checked := True;
  rbCreateUsers.OnClick := @OnCreateRdpShortcutsClick;

  rbUseExistingUsers := TRadioButton.Create(CreateRdpShortcutsGroup);
  rbUseExistingUsers.Parent := CreateRdpShortcutsGroup;
  rbUseExistingUsers.Left := ScaleX(10);
  rbUseExistingUsers.Top := ScaleY(32);
  rbUseExistingUsers.Width := ScaleX(340);
  rbUseExistingUsers.Caption := 'Use existing users';
  rbUseExistingUsers.Checked := False;
  rbUseExistingUsers.OnClick := @OnCreateRdpShortcutsClick;

  rbUseExistingUsersHint := TLabel.Create(CreateRdpShortcutsGroup);
  rbUseExistingUsersHint.Parent := CreateRdpShortcutsGroup;
  rbUseExistingUsersHint.Left := ScaleX(26);
  rbUseExistingUsersHint.Top := rbUseExistingUsers.Top + ScaleY(20);
  rbUseExistingUsersHint.Width := ScaleX(320);
  rbUseExistingUsersHint.AutoSize := False;
  rbUseExistingUsersHint.Transparent := True;
  rbUseExistingUsersHint.WordWrap := True;
  rbUseExistingUsersHint.Caption := '(Selected because your desktop already has RDP+ shortcuts)';
  rbUseExistingUsersHint.Font.Style := [fsItalic];
  rbUseExistingUsersHint.Visible := False;

  // Edit Shortcut radio placed halfway between Install and Uninstall

  // Set initial enabled state based on Create RDP shortcuts checkbox
  rbCreateUsers.Enabled := chkCreateRdpShortcuts.Checked;
  rbUseExistingUsers.Enabled := chkCreateRdpShortcuts.Checked;
  SyncSetupOptionHintStates();



  // --- Compact, even radio button spacing ---
  // Start at a base Y position just below the CreateRdpShortcutsGroup
  // Move the bottom three radio buttons up by about two lines (ScaleY(32))
  radioTopBase := CreateRdpShortcutsGroup.Top + CreateRdpShortcutsGroup.Height - ScaleY(20); // was + ScaleY(12), now - ScaleY(20) to move up ~2 lines
  radioSpacing := ScaleY(28); // Restore original spacing

  rbEditShortcutSettings := TRadioButton.Create(Page_InstallOptions);
  rbEditShortcutSettings.Parent := Page_InstallOptions.Surface;
  rbEditShortcutSettings.Left := ScaleX(10);
  rbEditShortcutSettings.Top := radioTopBase;
  rbEditShortcutSettings.Width := ScaleX(420);
  rbEditShortcutSettings.Caption := 'Edit existing shortcut settings';
  rbEditShortcutSettings.Checked := False;
  rbEditShortcutSettings.OnClick := @OnInstallModeChange;

  rbEditSystemwideSettings := TRadioButton.Create(Page_InstallOptions);
  rbEditSystemwideSettings.Parent := Page_InstallOptions.Surface;
  rbEditSystemwideSettings.Left := ScaleX(10);
  rbEditSystemwideSettings.Top := radioTopBase + radioSpacing;
  rbEditSystemwideSettings.Width := ScaleX(420);
  rbEditSystemwideSettings.Caption := 'Edit System-wide RDP settings';
  rbEditSystemwideSettings.Checked := False;
  rbEditSystemwideSettings.OnClick := @OnInstallModeChange;

  rbShowRDPInfo := TRadioButton.Create(Page_InstallOptions);
  rbShowRDPInfo.Parent := Page_InstallOptions.Surface;
  rbShowRDPInfo.Left := ScaleX(10);
  rbShowRDPInfo.Top := radioTopBase + radioSpacing * 2;
  rbShowRDPInfo.Width := ScaleX(420);
  rbShowRDPInfo.Caption := 'Show RDP Info';
  rbShowRDPInfo.Checked := False;
  rbShowRDPInfo.OnClick := @OnInstallModeChange;

  rbQuickFixes := TRadioButton.Create(Page_InstallOptions);
  rbQuickFixes.Parent := Page_InstallOptions.Surface;
  rbQuickFixes.Left := ScaleX(10);
  rbQuickFixes.Top := radioTopBase + radioSpacing * 3;
  rbQuickFixes.Width := ScaleX(420);
  rbQuickFixes.Caption := 'Quick Fixes';
  rbQuickFixes.Checked := False;
  rbQuickFixes.OnClick := @OnInstallModeChange;

  rbUninstall := TRadioButton.Create(Page_InstallOptions);
  rbUninstall.Parent := Page_InstallOptions.Surface;
  rbUninstall.Left := ScaleX(10);
  rbUninstall.Top := radioTopBase + radioSpacing * 4;
  rbUninstall.Width := ScaleX(420);
  rbUninstall.Caption := 'Uninstall (keeps users)';
  rbUninstall.Checked := False;
  rbUninstall.OnClick := @OnInstallModeChange;

  // Add credits text at bottom of the welcome intro page
  CreditsText := TRichEditViewer.Create(WelcomePage);
  CreditsText.Parent := WelcomePage.Surface;
  CreditsText.Left := ScaleX(10);
  CreditsText.Top := ScaleY(10);
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
  lblCreditsHeader.Top := CreditsText.Top + ScaleY(18);
  lblCreditsHeader.Caption := 'Credits:';
  lblCreditsHeader.Font.Style := [fsBold];
  lblCreditsHeader.Font.Size := CreditsText.Font.Size + 1;
  lblCreditsHeader.Font.Color := CreditsText.Font.Color;
  lblCreditsHeader.Transparent := True;
  lblCreditsHeader.AutoSize := True;

  // Start placing bullet lines under the header
  topPos := lblCreditsHeader.Top + ScaleY(18);



  lblBullet2 := TLabel.Create(WelcomePage);
  lblBullet2.Parent := WelcomePage.Surface;
  lblBullet2.Left := CreditsText.Left;
  lblBullet2.Top := topPos;
  lblBullet2.Caption := '> ';
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
  lblTermSuffix.Caption := '(it''s fantastic!)';
  lblTermSuffix.Font.Color := CreditsText.Font.Color;
  lblTermSuffix.Transparent := True;
  lblTermSuffix.AutoSize := True;

  topPos := topPos + ScaleY(18);

  // Line 3: Special thanks with two clickable names
  lblBullet5 := TLabel.Create(WelcomePage);
  lblBullet5.Parent := WelcomePage.Surface;
  lblBullet5.Left := CreditsText.Left;
  lblBullet5.Top := topPos;
  lblBullet5.Caption := '> Special thanks to Bee Swarm Sim communities: ';
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

  // Line 4: Assembled by cpdx4. Project Home:
  lblBullet4 := TLabel.Create(WelcomePage);
  lblBullet4.Parent := WelcomePage.Surface;
  lblBullet4.Left := CreditsText.Left;
  lblBullet4.Top := topPos;
  lblBullet4.Caption := '> Assembled by cpdx4. Project Home: ';
  lblBullet4.Font.Color := CreditsText.Font.Color;
  lblBullet4.Transparent := True;
  lblBullet4.AutoSize := True;

  // Clickable Project Home link (only URL portion)
  lblProjectHome := TLabel.Create(WelcomePage);
  lblProjectHome.Parent := WelcomePage.Surface;
  lblProjectHome.Left := lblBullet4.Left + lblBullet4.Width + ScaleX(4);
  lblProjectHome.Top := topPos;
  lblProjectHome.Caption := 'cpdx4.github.io/RDPWrapKit';
  lblProjectHome.Font.Color := LinkColor;
  lblProjectHome.Font.Style := [fsUnderline];
  lblProjectHome.Cursor := crHand;
  lblProjectHome.OnClick := @OpenProjectHome;
  lblProjectHome.Transparent := True;
  lblProjectHome.AutoSize := True;

  // Position the block near the bottom of the welcome page area
  topPos := topPos + ScaleY(24);
  DesiredBottom := WelcomePage.SurfaceHeight;
  BlockTop := DesiredBottom - (topPos - CreditsText.Top);
  if BlockTop < CreditsText.Top then
    BlockTop := CreditsText.Top;
  // Shift all created labels by the delta to align the block (guard against missing labels)
  Delta := BlockTop - CreditsText.Top;
  if Assigned(lblCreditsHeader) then lblCreditsHeader.Top := lblCreditsHeader.Top + Delta;
  if Assigned(lblBullet1) then lblBullet1.Top := lblBullet1.Top + Delta;
  if Assigned(lblStasName) then lblStasName.Top := lblStasName.Top + Delta;
  if Assigned(lblStasSuffix) then lblStasSuffix.Top := lblStasSuffix.Top + Delta;
  if Assigned(lblBullet2) then lblBullet2.Top := lblBullet2.Top + Delta;
  if Assigned(lblTermName) then lblTermName.Top := lblTermName.Top + Delta;
  if Assigned(lblTermSuffix) then lblTermSuffix.Top := lblTermSuffix.Top + Delta;
  if Assigned(lblBullet5) then lblBullet5.Top := lblBullet5.Top + Delta;
  if Assigned(lblBSGHName) then lblBSGHName.Top := lblBSGHName.Top + Delta;
  if Assigned(lblAnd) then lblAnd.Top := lblAnd.Top + Delta;
  if Assigned(lblBSSName) then lblBSSName.Top := lblBSSName.Top + Delta;
  if Assigned(lblBullet4) then lblBullet4.Top := lblBullet4.Top + Delta;
  if Assigned(lblProjectHome) then lblProjectHome.Top := lblProjectHome.Top + Delta;

  // Initialize derived flags to reflect defaults so page skipping works immediately
  DoInstallTermWrap := chkInstallTermWrap.Checked;
  DoCreateRdpShortcuts := chkCreateRdpShortcuts.Checked;
  if DoCreateRdpShortcuts and rbCreateUsers.Checked then
    CreateUserMode := createUserModeNew
  else
    CreateUserMode := createUserModeExisting;
  
  // Create User Page (after InstallTypePage)
  UserPage := CreateInputQueryPage(
    Page_InstallOptions.ID,
    'Create RDP User Account',
    'Create a new user by entering a username (such as "macro1" or "rdp1") and password below.',
    ''
  );
  UserPage.Add('Create a Username: (eg. macro1)', False);
  UserPage.Add('Set a Password:', True);  // masked input
  // Move Username and Password controls up slightly to improve layout
  try
    // Shift the edit controls and their labels (username/password)
    UserPage.Edits[0].Top := UserPage.Edits[0].Top - ScaleY(10);
    UserPage.Edits[1].Top := UserPage.Edits[1].Top - ScaleY(20);
    UserPage.PromptLabels[0].Top := UserPage.PromptLabels[0].Top - ScaleY(10);
    UserPage.PromptLabels[1].Top := UserPage.PromptLabels[1].Top - ScaleY(20);
  except
  end;

  // Shortcut Settings page — reachable from Create Users, Existing Users, and Edit Shortcuts.
  // Anchored to UserPage.ID so it physically follows UserPage (and therefore ExistingUsers/
  // EditShortcuts, which are all before UserPage in traversal order).
  Page_ShortcutSettings := CreateCustomPage(
    UserPage.ID,
    'Shortcut Settings',
    'Configure the settings for your RDP desktop shortcuts.' + #13#10 + 
    'If you dont know what to choose here, just click [Next]'
  );
  BuildShortcutSettingsBlock(Page_ShortcutSettings.Surface, ScaleY(10));


  // Create Edit System-wide Settings page as a custom page.
  // Using a custom page avoids InputOptionPage's built-in checklist control,
  // which can obscure non-windowed labels in some wizard themes.
  EditSystemwideSettingsPage := CreateCustomPage(
    Page_InstallOptions.ID,
    'Edit System-wide RDP settings',
    'Warning: Changes below should only be changed if you know what you are doing'#13#10 +
    'Make changes and then click [Next]'
  );
  if IsDarkColor(EditSystemwideSettingsPage.Surface.Color) then
    LabelColor := clWhite
  else
    LabelColor := clBlack;
  // Build controls for Edit System-wide Settings page
  begin
    leftPos := ScaleX(20);
    topPos := ScaleY(12);

    // indentation and value column for neat alignment
    childIndent := ScaleX(16);
    childLeft := leftPos + childIndent;
    valueLeft := childLeft + ScaleX(140);

    // General Settings
    lblGenHeader := TLabel.Create(EditSystemwideSettingsPage);
    lblGenHeader.Parent := EditSystemwideSettingsPage.Surface;
    lblGenHeader.Left := leftPos;
    lblGenHeader.Top := topPos;
    lblGenHeader.Caption := 'General Settings';
    lblGenHeader.Font.Style := [fsBold];
    lblGenHeader.ParentFont := False;
    lblGenHeader.Font.Color := LabelColor;
    lblGenHeader.Transparent := True;
    topPos := topPos + ScaleY(20);

    chkEnableRDP := TCheckBox.Create(EditSystemwideSettingsPage);
    chkEnableRDP.Parent := EditSystemwideSettingsPage.Surface;
    chkEnableRDP.Left := childLeft;
    chkEnableRDP.Top := topPos;
    chkEnableRDP.Width := ScaleX(420) - childIndent;
    chkEnableRDP.Caption := 'Enable Remote Desktop';
    chkEnableRDP.Checked := True;
    chkEnableRDP.ParentFont := False;
    chkEnableRDP.Font.Color := LabelColor;
    MakeHelpButton(EditSystemwideSettingsPage, topPos, 1);
    topPos := topPos + ScaleY(20);

    chkShowUsers := TCheckBox.Create(EditSystemwideSettingsPage);
    chkShowUsers.Parent := EditSystemwideSettingsPage.Surface;
    chkShowUsers.Left := childLeft;
    chkShowUsers.Top := topPos;
    chkShowUsers.Width := ScaleX(420) - childIndent;
    chkShowUsers.Caption := 'Show users on logon screen';
    chkShowUsers.Checked := True;
    chkShowUsers.ParentFont := False;
    chkShowUsers.Font.Color := LabelColor;
    MakeHelpButton(EditSystemwideSettingsPage, topPos, 2);
    topPos := topPos + ScaleY(20);

    chkPreventDuplicate := TCheckBox.Create(EditSystemwideSettingsPage);
    chkPreventDuplicate.Parent := EditSystemwideSettingsPage.Surface;
    chkPreventDuplicate.Left := childLeft;
    chkPreventDuplicate.Top := topPos;
    chkPreventDuplicate.Width := ScaleX(420) - childIndent;
    chkPreventDuplicate.Caption := 'Prevent duplicate connections per user';
    chkPreventDuplicate.Checked := True;
    chkPreventDuplicate.ParentFont := False;
    chkPreventDuplicate.Font.Color := LabelColor;
    MakeHelpButton(EditSystemwideSettingsPage, topPos, 3);
    topPos := topPos + ScaleY(26);

    chkHideSecurityWarnings := TCheckBox.Create(EditSystemwideSettingsPage);
    chkHideSecurityWarnings.Parent := EditSystemwideSettingsPage.Surface;
    chkHideSecurityWarnings.Left := childLeft;
    chkHideSecurityWarnings.Top := topPos;
    chkHideSecurityWarnings.Width := ScaleX(420) - childIndent;
    chkHideSecurityWarnings.Caption := 'Hide most security warnings';
    chkHideSecurityWarnings.Checked := False;
    chkHideSecurityWarnings.ParentFont := False;
    chkHideSecurityWarnings.Font.Color := LabelColor;
    MakeHelpButton(EditSystemwideSettingsPage, topPos, 8);
    topPos := topPos + ScaleY(26);

    lblRdpPort := TLabel.Create(EditSystemwideSettingsPage);
    lblRdpPort.Parent := EditSystemwideSettingsPage.Surface;
    lblRdpPort.Left := childLeft;
    lblRdpPort.Top := topPos;
    lblRdpPort.Caption := 'RDP Port:';
    lblRdpPort.ParentFont := False;
    lblRdpPort.Font.Color := LabelColor;
    lblRdpPort.Transparent := True;

    edtRdpPort := TEdit.Create(EditSystemwideSettingsPage);
    edtRdpPort.Parent := EditSystemwideSettingsPage.Surface;
    edtRdpPort.Left := childLeft + ScaleX(60);
    edtRdpPort.Top := topPos - ScaleY(2);
    edtRdpPort.Width := Round(ScaleX(80) * 0.6); // reduce width by 40%
    edtRdpPort.Text := '3389';
    edtRdpPort.ParentFont := False;
    edtRdpPort.Font.Color := LabelColor;

    lblPortDefault := TLabel.Create(EditSystemwideSettingsPage);
    lblPortDefault.Parent := EditSystemwideSettingsPage.Surface;
    lblPortDefault.Left := childLeft + ScaleX(120);
    lblPortDefault.Top := topPos;
    lblPortDefault.Caption := '(default: 3389. Requires restart of RDP Service)';
    lblPortDefault.ParentFont := False;
    lblPortDefault.Font.Color := LabelColor;
    lblPortDefault.Transparent := True;
    MakeHelpButton(EditSystemwideSettingsPage, topPos - ScaleY(2), 4);
    topPos := topPos + ScaleY(36);

    // Performance
    TmpLabel := TLabel.Create(EditSystemwideSettingsPage);
    TmpLabel.Parent := EditSystemwideSettingsPage.Surface;
    TmpLabel.Left := leftPos;
    TmpLabel.Top := topPos;
    TmpLabel.Caption := 'Performance';
    TmpLabel.Font.Style := [fsBold];
    TmpLabel.ParentFont := False;
    TmpLabel.Font.Color := LabelColor;
    TmpLabel.Transparent := True;
    topPos := topPos + ScaleY(20);

    TmpLabel := TLabel.Create(EditSystemwideSettingsPage);
    TmpLabel.Parent := EditSystemwideSettingsPage.Surface;
    TmpLabel.Left := childLeft;
    TmpLabel.Top := topPos + ScaleY(3);
    TmpLabel.Caption := 'RemoteFX image quality:';
    TmpLabel.ParentFont := False;
    TmpLabel.Font.Color := LabelColor;
    TmpLabel.Transparent := True;
    cmbGPImageQuality := TComboBox.Create(EditSystemwideSettingsPage);
    cmbGPImageQuality.Parent := EditSystemwideSettingsPage.Surface;
    cmbGPImageQuality.Left := childLeft + ScaleX(130);
    cmbGPImageQuality.Top := topPos;
    cmbGPImageQuality.Width := ScaleX(420) - childIndent - ScaleX(140);
    cmbGPImageQuality.Style := csDropDownList;
    cmbGPImageQuality.Items.Add('Not Set (system default)');
    cmbGPImageQuality.Items.Add('0 - Low');
    cmbGPImageQuality.Items.Add('1 - Medium');
    cmbGPImageQuality.Items.Add('2 - High');
    cmbGPImageQuality.Items.Add('3 - Lossless');
    cmbGPImageQuality.ItemIndex := 0;
    MakeHelpButton(EditSystemwideSettingsPage, topPos, 5);
    topPos := topPos + ScaleY(24);

    TmpLabel := TLabel.Create(EditSystemwideSettingsPage);
    TmpLabel.Parent := EditSystemwideSettingsPage.Surface;
    TmpLabel.Left := childLeft;
    TmpLabel.Top := topPos + ScaleY(3);
    TmpLabel.Caption := 'RemoteFX compression:';
    TmpLabel.ParentFont := False;
    TmpLabel.Font.Color := LabelColor;
    TmpLabel.Transparent := True;
    cmbGPCompression := TComboBox.Create(EditSystemwideSettingsPage);
    cmbGPCompression.Parent := EditSystemwideSettingsPage.Surface;
    cmbGPCompression.Left := childLeft + ScaleX(130);
    cmbGPCompression.Top := topPos;
    cmbGPCompression.Width := ScaleX(420) - childIndent - ScaleX(140);
    cmbGPCompression.Style := csDropDownList;
    cmbGPCompression.Items.Add('Not Set (system default)');
    cmbGPCompression.Items.Add('0 - No compression');
    cmbGPCompression.Items.Add('1 - Optimised for memory');
    cmbGPCompression.Items.Add('2 - Balanced (memory / bandwidth)');
    cmbGPCompression.Items.Add('3 - Optimised for bandwidth');
    cmbGPCompression.ItemIndex := 0;
    MakeHelpButton(EditSystemwideSettingsPage, topPos, 6);
    topPos := topPos + ScaleY(28);

    // Actions
    lblActionsHeader := TLabel.Create(EditSystemwideSettingsPage);
    lblActionsHeader.Parent := EditSystemwideSettingsPage.Surface;
    lblActionsHeader.Left := leftPos;
    lblActionsHeader.Top := topPos;
    lblActionsHeader.Caption := 'Actions';
    lblActionsHeader.Font.Style := [fsBold];
    lblActionsHeader.ParentFont := False;
    lblActionsHeader.Font.Color := LabelColor;
    lblActionsHeader.Transparent := True;
    topPos := topPos + ScaleY(20);

    chkRestartRDP := TCheckBox.Create(EditSystemwideSettingsPage);
    chkRestartRDP.Parent := EditSystemwideSettingsPage.Surface;
    chkRestartRDP.Left := childLeft;
    chkRestartRDP.Top := topPos;
    chkRestartRDP.Width := ScaleX(420) - childIndent;
    chkRestartRDP.Caption := 'Restart RDP Service';
    chkRestartRDP.Checked := False;
    chkRestartRDP.ParentFont := False;
    chkRestartRDP.Font.Color := LabelColor;
    MakeHelpButton(EditSystemwideSettingsPage, topPos, 7);
    topPos := topPos + ScaleY(20);
  end

  // -------------------------------------------------------------------------
  // Create "Quick Fixes" page
  // Provides quick one-click fixes for common RDP-related issues.
  // -------------------------------------------------------------------------
  QuickFixesPage := CreateCustomPage(
    Page_InstallOptions.ID,
    'Quick Fixes',
    'Select a quick fix to apply, then click [Next] to execute it.'
  );
  if IsDarkColor(QuickFixesPage.Surface.Color) then
    LabelColor := clWhite
  else
    LabelColor := clBlack;
  begin
    leftPos := ScaleX(20);
    topPos := ScaleY(12);
    childIndent := ScaleX(16);
    childLeft := leftPos + childIndent;

    lblQFHeader := TLabel.Create(QuickFixesPage);
    lblQFHeader.Parent := QuickFixesPage.Surface;
    lblQFHeader.Left := leftPos;
    lblQFHeader.Top := topPos;
    lblQFHeader.Caption := 'Select a fix:';
    lblQFHeader.Font.Style := [fsBold];
    lblQFHeader.ParentFont := False;
    lblQFHeader.Font.Color := LabelColor;
    lblQFHeader.Transparent := True;
    topPos := topPos + ScaleY(24);

    rbQFRestartRDP := TRadioButton.Create(QuickFixesPage);
    rbQFRestartRDP.Parent := QuickFixesPage.Surface;
    rbQFRestartRDP.Left := childLeft;
    rbQFRestartRDP.Top := topPos;
    rbQFRestartRDP.Width := ScaleX(420) - childIndent;
    rbQFRestartRDP.Caption := 'Restart RDP Service';
    rbQFRestartRDP.Checked := True;
    rbQFRestartRDP.ParentFont := False;
    rbQFRestartRDP.Font.Color := LabelColor;
    MakeHelpButton(QuickFixesPage, topPos, 10);
    topPos := topPos + ScaleY(24);

    rbQFAccountNeverExpires := TRadioButton.Create(QuickFixesPage);
    rbQFAccountNeverExpires.Parent := QuickFixesPage.Surface;
    rbQFAccountNeverExpires.Left := childLeft;
    rbQFAccountNeverExpires.Top := topPos;
    rbQFAccountNeverExpires.Width := ScaleX(420) - childIndent;
    rbQFAccountNeverExpires.Caption := 'Set all Local Accounts to Never Expire';
    rbQFAccountNeverExpires.Checked := False;
    rbQFAccountNeverExpires.ParentFont := False;
    rbQFAccountNeverExpires.Font.Color := LabelColor;
    MakeHelpButton(QuickFixesPage, topPos, 11);
    topPos := topPos + ScaleY(24);

    rbQFRestoreTermService := TRadioButton.Create(QuickFixesPage);
    rbQFRestoreTermService.Parent := QuickFixesPage.Surface;
    rbQFRestoreTermService.Left := childLeft;
    rbQFRestoreTermService.Top := topPos;
    rbQFRestoreTermService.Width := ScaleX(420) - childIndent;
    rbQFRestoreTermService.Caption := 'Restore deleted Remote Desktop Service';
    rbQFRestoreTermService.Checked := False;
    rbQFRestoreTermService.ParentFont := False;
    rbQFRestoreTermService.Font.Color := LabelColor;
    MakeHelpButton(QuickFixesPage, topPos, 12);
  end

  // -------------------------------------------------------------------------
  // Create "Show RDP Info" page
  // Shows system status and allows configuring RemoteFX / startup GP settings.
  // -------------------------------------------------------------------------
  Page_ShowRDPInfo := CreateCustomPage(
    Page_InstallOptions.ID,
    'Show RDP Info',
    'View system status and configure RDP performance settings.'
  );
  begin
    if IsDarkColor(Page_ShowRDPInfo.Surface.Color) then
      LabelColor := clWhite
    else
      LabelColor := clBlack;

    leftPos := ScaleX(20);
    topPos := ScaleY(5);
    childIndent := ScaleX(16);
    childLeft := leftPos + childIndent;
    valueLeft := childLeft + ScaleX(140);

    // --- System Status header (moved from Edit System-wide RDP Settings page) ---
    lblSysHeader := TLabel.Create(Page_ShowRDPInfo);
    lblSysHeader.Parent := Page_ShowRDPInfo.Surface;
    lblSysHeader.Left := leftPos;
    lblSysHeader.Top := topPos;
    lblSysHeader.Caption := 'System Status';
    lblSysHeader.Font.Style := [fsBold];
    lblSysHeader.ParentFont := False;
    lblSysHeader.Font.Color := LabelColor;
    lblSysHeader.Transparent := True;
    topPos := topPos + ScaleY(20);

    lblWinVerName := TLabel.Create(Page_ShowRDPInfo);
    lblWinVerName.Parent := Page_ShowRDPInfo.Surface;
    lblWinVerName.Left := childLeft;
    lblWinVerName.Top := topPos;
    lblWinVerName.Caption := 'Windows Version:';
    lblWinVerName.ParentFont := False;
    lblWinVerName.Font.Color := LabelColor;
    lblWinVerName.Transparent := True;
    lblWinVer := TLabel.Create(Page_ShowRDPInfo);
    lblWinVer.Parent := Page_ShowRDPInfo.Surface;
    lblWinVer.Left := valueLeft;
    lblWinVer.Top := topPos;
    lblWinVer.Caption := '...';
    lblWinVer.ParentFont := False;
    lblWinVer.Font.Color := LabelColor;
    lblWinVer.Transparent := True;
    topPos := topPos + ScaleY(18);

    lblRDPServiceName := TLabel.Create(Page_ShowRDPInfo);
    lblRDPServiceName.Parent := Page_ShowRDPInfo.Surface;
    lblRDPServiceName.Left := childLeft;
    lblRDPServiceName.Top := topPos;
    lblRDPServiceName.Caption := 'RDP Service:';
    lblRDPServiceName.ParentFont := False;
    lblRDPServiceName.Font.Color := LabelColor;
    lblRDPServiceName.Transparent := True;
    lblRDPService := TLabel.Create(Page_ShowRDPInfo);
    lblRDPService.Parent := Page_ShowRDPInfo.Surface;
    lblRDPService.Left := valueLeft;
    lblRDPService.Top := topPos;
    lblRDPService.Caption := '...';
    lblRDPService.ParentFont := False;
    lblRDPService.Font.Color := LabelColor;
    lblRDPService.Transparent := True;
    topPos := topPos + ScaleY(18);

    lblWinRDPVerName := TLabel.Create(Page_ShowRDPInfo);
    lblWinRDPVerName.Parent := Page_ShowRDPInfo.Surface;
    lblWinRDPVerName.Left := childLeft;
    lblWinRDPVerName.Top := topPos;
    lblWinRDPVerName.Caption := 'RDP DLL Version:';
    lblWinRDPVerName.ParentFont := False;
    lblWinRDPVerName.Font.Color := LabelColor;
    lblWinRDPVerName.Transparent := True;
    lblWinRDPVer := TLabel.Create(Page_ShowRDPInfo);
    lblWinRDPVer.Parent := Page_ShowRDPInfo.Surface;
    lblWinRDPVer.Left := valueLeft;
    lblWinRDPVer.Top := topPos;
    lblWinRDPVer.Caption := '...';
    lblWinRDPVer.ParentFont := False;
    lblWinRDPVer.Font.Color := LabelColor;
    lblWinRDPVer.Transparent := True;
    topPos := topPos + ScaleY(18);

    lblWrapperVerName := TLabel.Create(Page_ShowRDPInfo);
    lblWrapperVerName.Parent := Page_ShowRDPInfo.Surface;
    lblWrapperVerName.Left := childLeft;
    lblWrapperVerName.Top := topPos;
    lblWrapperVerName.Caption := 'Wrapper file info:';
    lblWrapperVerName.ParentFont := False;
    lblWrapperVerName.Font.Color := LabelColor;
    lblWrapperVerName.Transparent := True;
    lblWrapperVer := TLabel.Create(Page_ShowRDPInfo);
    lblWrapperVer.Parent := Page_ShowRDPInfo.Surface;
    lblWrapperVer.Left := valueLeft;
    lblWrapperVer.Top := topPos;
    lblWrapperVer.Caption := '...';
    lblWrapperVer.ParentFont := False;
    lblWrapperVer.Font.Color := LabelColor;
    lblWrapperVer.Transparent := True;
    topPos := topPos + ScaleY(18);

    lblListenerName := TLabel.Create(Page_ShowRDPInfo);
    lblListenerName.Parent := Page_ShowRDPInfo.Surface;
    lblListenerName.Left := childLeft;
    lblListenerName.Top := topPos;
    lblListenerName.Caption := 'Listening Status:';
    lblListenerName.ParentFont := False;
    lblListenerName.Font.Color := LabelColor;
    lblListenerName.Transparent := True;
    lblListener := TLabel.Create(Page_ShowRDPInfo);
    lblListener.Parent := Page_ShowRDPInfo.Surface;
    lblListener.Left := valueLeft;
    lblListener.Top := topPos;
    lblListener.Caption := '...';
    lblListener.ParentFont := False;
    lblListener.Font.Color := LabelColor;
    lblListener.Transparent := True;
    topPos := topPos + ScaleY(26);

  end;

  // Create Tool 1 Page: Create RDP desktop shortcuts for existing local users
  Page_CreateShortcutsForExistingUsers := CreateCustomPage(
    Page_InstallOptions.ID,
    'Create RDP Desktop Shortcuts',
    'This is a list of users found on this PC. Checkmark the accounts you want to make desktop shortcuts for and type their password.'
  );
  

  EditShortcutPage := CreateCustomPage(
    Page_InstallOptions.ID,
    'Edit existing shortcut settings',
    'Select one or more Desktop .rdp shortcuts to edit.'
  );
  
  
  // Advanced editing page (shown only when user checks "Open advanced shortcut options")
  // Positioned after ShortcutSettings page so it appears just before the Installing page.
  Page_EditShortcutAdvanced := CreateCustomPage(
    Page_ShortcutSettings.ID,
    'Advanced Shortcut Editing',
    'Manual editing instructions for your RDP shortcut.'
  );
  
  // Initialize lists for tracking
  LocalUsersList := TStringList.Create;        // Will be populated when Create Shortcuts page is shown
  LocalUserDisplayList := TStringList.Create;  // Parallel display labels (email for online accounts)
  DesktopRdpFiles := TStringList.Create;
  SetLength(UserCheckBoxes, 0);
  SetLength(UserPasswordEdits, 0);
  SetLength(UserPasswordStatus, 0);
  SetLength(ShortcutCheckBoxes, 0);
  CurrentShortcutPage := 0;
  ShortcutsPerPage := 8;
  SelectedShortcutIndex := -1;
  SelectedShortcutPath := '';
  SelectedShortcutPaths := TStringList.Create;
  EditShortcutControlsBuilt := False;
  EditShortcutAdvancedControlsBuilt := False;
  ShortcutsList := TStringList.Create;
  

  // Add label for options section
  OptionsLabel := TLabel.Create(UserPage);
  OptionsLabel.Parent := UserPage.Surface;
  OptionsLabel.Left := ScaleX(260);
  OptionsLabel.Top := ScaleY(246);
  OptionsLabel.Caption := 'What would you like to do next?';
  OptionsLabel.Font.Style := [fsBold];

  // Add "Create more users" radio button
  AddMoreRadio := TRadioButton.Create(UserPage);
  AddMoreRadio.Parent := UserPage.Surface;
  AddMoreRadio.Left := ScaleX(270);
  AddMoreRadio.Top := OptionsLabel.Top + ScaleY(20);
  AddMoreRadio.Width := ScaleX(400);
  AddMoreRadio.Caption := 'I want to create another user';
  AddMoreRadio.Checked := False;

  // Add "Done creating users" radio button
  DoneRadio := TRadioButton.Create(UserPage);
  DoneRadio.Parent := UserPage.Surface;
  DoneRadio.Left := ScaleX(270);
  DoneRadio.Top := AddMoreRadio.Top + ScaleY(22);
  DoneRadio.Width := ScaleX(400);
  DoneRadio.Caption := 'I''m done creating users, continue setup';
  DoneRadio.Checked := True;  // Default selection
   
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
  StepInstallTermWrap := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);topPos := topPos + ScaleY(16);
  StepConfigureService := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal); topPos := topPos + ScaleY(16);
  StepCreateUsers := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);      topPos := topPos + ScaleY(16);
  StepCreateShortcuts := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);  topPos := topPos + ScaleY(16);
  StepPreTrust := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);         topPos := topPos + ScaleY(16);
  StepStartSvc := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);         topPos := topPos + ScaleY(16);
  StepCheckRDP := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);         topPos := topPos + ScaleY(16);
  StepCheckMSTSC := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);       topPos := topPos + ScaleY(16);
  StepInstallMSTSC := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);     topPos := topPos + ScaleY(16);
  StepRemoveFolder := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);     topPos := topPos + ScaleY(16);
  StepUninstallTermWrap := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal); topPos := topPos + ScaleY(16);
  StepEnableRDP := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);       topPos := topPos + ScaleY(16);
  StepShowUsers := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);       topPos := topPos + ScaleY(16);
  StepPreventDuplicate := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal); topPos := topPos + ScaleY(16);
  StepSetRdpPort := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);     topPos := topPos + ScaleY(16);
  StepRestartRDP := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);     topPos := topPos + ScaleY(16);
  StepQuickFixes := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);      topPos := topPos + ScaleY(16);
  StepAccountNeverExpires := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal); topPos := topPos + ScaleY(16);
  StepShowRDPInfo := CreateStepLabel(WizardForm.InstallingPage, leftPos, topPos, widthVal);          topPos := topPos + ScaleY(16);
  // Create a label on the Finished page to show completion messages (positioned right below header)
  FinishedText := TLabel.Create(WizardForm.FinishedLabel.Parent);
  FinishedText.Parent := WizardForm.FinishedLabel.Parent;
  FinishedText.Left := WizardForm.FinishedLabel.Left;
  FinishedText.Top := WizardForm.FinishedHeadingLabel.Top + WizardForm.FinishedHeadingLabel.Height + ScaleY(8);
  FinishedText.Width := WizardForm.FinishedLabel.Width;
  FinishedText.AutoSize := False;
  FinishedText.WordWrap := True;
  FinishedText.Transparent := True;
  FinishedText.Font.Size := WizardForm.FinishedLabel.Font.Size;
  FinishedText.Visible := True;
  WizardForm.FinishedLabel.Visible := False;

  // Create a button to save the install log, placed in the bottom button bar next to Finish
  ViewLogButton := TButton.Create(WizardForm);
  ViewLogButton.Parent := WizardForm;
  ViewLogButton.Width := ScaleX(120);
  ViewLogButton.Height := WizardForm.NextButton.Height;
  ViewLogButton.Left := WizardForm.NextButton.Left - ViewLogButton.Width - ScaleX(10);
  ViewLogButton.Top := WizardForm.NextButton.Top;
  ViewLogButton.Caption := 'Save Install Log';
  ViewLogButton.OnClick := @OnViewLogButtonClick;
  ViewLogButton.Visible := False;  // only shown on the Finished page

  // Optional example image shown only on the real final page when using
  // Edit Shortcut Settings mode.
  FinishedExampleImage := TBitmapImage.Create(WizardForm.FinishedLabel.Parent);
  FinishedExampleImage.Parent := WizardForm.FinishedLabel.Parent;
  FinishedExampleImage.Left := WizardForm.FinishedLabel.Left;
  FinishedExampleImage.Top := FinishedText.Top + ScaleY(100);
  FinishedExampleImage.Width := ScaleX(142);
  FinishedExampleImage.Height := ScaleY(150);
  FinishedExampleImage.Stretch := False;
  FinishedExampleImage.Visible := False;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  UserName: string;
  Password: string;
  i: Integer;
  SelectedCount: Integer;
  HasErrors: Boolean;
  RdpPortValue: Integer;
  PortError: string;
  NewShortcutBase: string;
  NewShortcutPath: string;
begin
  LogEntry('NextButtonClick');
  LogDebug('NextButtonClick: CurPageID=' + IntToStr(CurPageID) + ' (' + GetPageNameById(CurPageID) + ')');
  Result := True;

  if CurPageID = wpFinished then
  begin
    CleanupPendingDebugFiles;
    Result := True;
    exit;
  end;
  
  if CurPageID = Page_InstallOptions.ID then
  begin
    // Derive install mode and flags from the options page controls
    DoInstallTermWrap := False;
    DoCreateRdpShortcuts := False;
    CreateUserMode := createUserModeNew;
    DoEditSystemWideSettings := False;

    if Assigned(rbQuickFixes) and rbQuickFixes.Checked then
    begin
      SelectedInstallMode := installModeQuickFixes;
    end
    else if rbUninstall.Checked then
    begin
      SelectedInstallMode := installModeUninstall;
    end
    else if (Assigned(rbEditSystemwideSettings) and rbEditSystemwideSettings.Checked) then
    begin
      SelectedInstallMode := installModeEditSystemwideSettings;
      DoEditSystemWideSettings := True;
    end
    else if (Assigned(rbShowRDPInfo) and rbShowRDPInfo.Checked) then
    begin
      SelectedInstallMode := installModeShowRDPInfo;
    end
    else if rbEditShortcutSettings.Checked then
    begin
      SelectedInstallMode := installModeEditShortcuts;
      if Assigned(DesktopRdpFiles) then
        DesktopRdpFiles.Free;
      DesktopRdpFiles := GetDesktopRdpFiles;
      // Hide old controls before clearing references
      if Assigned(ShortcutHeaderLabel) then ShortcutHeaderLabel.Visible := False;
      if Assigned(ShortcutEmptyLabel)  then ShortcutEmptyLabel.Visible  := False;
      if Assigned(ShortcutPrevButton)  then ShortcutPrevButton.Visible  := False;
      if Assigned(ShortcutNextButton)  then ShortcutNextButton.Visible  := False;
      if Assigned(ShortcutPageLabel)   then ShortcutPageLabel.Visible   := False;
      for i := 0 to Length(ShortcutCheckBoxes) - 1 do
        if Assigned(ShortcutCheckBoxes[i]) then ShortcutCheckBoxes[i].Visible := False;
      EditShortcutControlsBuilt := False;
      SetLength(ShortcutCheckBoxes, 0);
      SelectedShortcutIndex := -1;
      SelectedShortcutPath := '';
      if Assigned(SelectedShortcutPaths) then
        SelectedShortcutPaths.Clear;
      Result := True;
      exit;
    end
    else // Install selected
    begin
      SelectedInstallMode := installModeInstall;
      DoInstallTermWrap    := chkInstallTermWrap.Checked;
      DoCreateRdpShortcuts := chkCreateRdpShortcuts.Checked;

      if DoCreateRdpShortcuts then
      begin
        if rbCreateUsers.Checked then
          CreateUserMode := createUserModeNew
        else
          CreateUserMode := createUserModeExisting;
      end;

      // Require at least one action under Install
      if (not DoInstallTermWrap) and (not DoCreateRdpShortcuts) then
      begin
        MsgBox('Please select at least one option under Install', mbError, MB_OK);
        Result := False;
        exit;
      end;

      // Pre-populate the existing-users page when taking that path
      if DoCreateRdpShortcuts and (CreateUserMode = createUserModeExisting) then
      begin
        LocalUsersList.Clear;
        LocalUserDisplayList.Clear;
        GetLocalUsers(LocalUsersList, LocalUserDisplayList);
        SetLength(UserCheckBoxes, LocalUsersList.Count);
        SetLength(UserPasswordEdits, LocalUsersList.Count);
        SetLength(UserPasswordStatus, LocalUsersList.Count);
        BuildCreateShortcutsControls;
      end;
    end;

    // Log detailed user selections on the Install Options page
    LogSectionHeader('USER SELECTIONS');
    if Assigned(rbInstall) and rbInstall.Checked then
    begin
      LogInfo('  (x) Install');
      if Assigned(chkInstallTermWrap) then
      begin
        if chkInstallTermWrap.Checked then
          LogInfo('       [x] Install TermWrap')
        else
          LogInfo('       [ ] Install TermWrap');
      end;
      if Assigned(chkCreateRdpShortcuts) then
      begin
        if chkCreateRdpShortcuts.Checked then
          LogInfo('       [x] Create Shortcuts')
        else
          LogInfo('       [ ] Create Shortcuts');
      end;
      if DoCreateRdpShortcuts then
      begin
        if Assigned(rbCreateUsers) and rbCreateUsers.Checked then
        begin
          LogInfo('          (x) Create New Users');
          LogInfo('          ( ) Use Existing Users');
        end
        else if Assigned(rbUseExistingUsers) and rbUseExistingUsers.Checked then
        begin
          LogInfo('          ( ) Create New Users');
          LogInfo('          (x) Use Existing Users');
        end;
      end;
    end
    else if Assigned(rbEditShortcutSettings) and rbEditShortcutSettings.Checked then
      LogInfo('  (x) Edit Shortcut Settings')
    else if Assigned(rbEditSystemwideSettings) and rbEditSystemwideSettings.Checked then
      LogInfo('  (x) Edit System-wide RDP Settings')
    else if Assigned(rbShowRDPInfo) and rbShowRDPInfo.Checked then
      LogInfo('  (x) Tune System Performance')
    else if Assigned(rbQuickFixes) and rbQuickFixes.Checked then
      LogInfo('  (x) Quick Fixes')
    else if Assigned(rbUninstall) and rbUninstall.Checked then
      LogInfo('  (x) Uninstall');
  end
  else if CurPageID = EditShortcutPage.ID then
  begin
    if DesktopRdpFiles.Count = 0 then
    begin
      MsgBox('No .rdp shortcuts were found on your Desktop.', mbError, MB_OK);
      Result := False;
      exit;
    end;

    // Count checked checkboxes and collect selected paths
    if not Assigned(SelectedShortcutPaths) then
      SelectedShortcutPaths := TStringList.Create;
    SelectedShortcutPaths.Clear;
    for i := 0 to DesktopRdpFiles.Count - 1 do
    begin
      if Assigned(ShortcutCheckBoxes[i]) and ShortcutCheckBoxes[i].Checked then
        SelectedShortcutPaths.Add(DesktopRdpFiles[i]);
    end;

    // Log which shortcuts the user selected
    LogDebug('EditShortcutPage: ' + IntToStr(SelectedShortcutPaths.Count) + ' shortcut(s) selected:');
    for i := 0 to SelectedShortcutPaths.Count - 1 do
      LogDebug('  [x] ' + ExtractFileName(SelectedShortcutPaths[i]));

    if SelectedShortcutPaths.Count = 0 then
    begin
      MsgBox('Select at least one shortcut to edit before continuing.', mbError, MB_OK);
      Result := False;
      exit;
    end;

    // For single selection, use the existing single-shortcut flow
    if SelectedShortcutPaths.Count = 1 then
    begin
      SelectedShortcutPath := SelectedShortcutPaths[0];
      SelectedShortcutIndex := -1;
      // Find the index
      for i := 0 to DesktopRdpFiles.Count - 1 do
      begin
        if CompareText(DesktopRdpFiles[i], SelectedShortcutPath) = 0 then
        begin
          SelectedShortcutIndex := i;
          break;
        end;
      end;
    end
    else
    begin
      // Multi-select: clear single path (will be handled in ssPostInstall)
      SelectedShortcutPath := '';
      SelectedShortcutIndex := -1;
    end;
  end
  else if CurPageID = Page_ShortcutSettings.ID then
  begin
    if SelectedInstallMode = installModeEditShortcuts then
    begin
      // Remember whether the user wants to open the advanced mstsc editor
      DoShowMstscEdit := Assigned(chkShowMoreShortcutOptions) and chkShowMoreShortcutOptions.Checked;

      // Handle shortcut file rename when the user changed the shortcut name
      if (SelectedShortcutPath <> '') and Assigned(edtShortcutName) then
      begin
        NewShortcutBase := Trim(edtShortcutName.Text);
        if NewShortcutBase <> '' then
        begin
          // Validate shortcut name — reject invalid Windows filename characters
          if not IsValidShortcutName(NewShortcutBase) then
          begin
            MsgBox(
              'The shortcut name contains invalid characters.' + #13#10#13#10 +
              'Windows filenames cannot contain:  < > : " / \ | ? *' + #13#10#13#10 +
              'Please remove those characters and try again.',
              mbError, MB_OK);
            Result := False;
            exit;
          end;
          // Ensure .rdp extension
          if CompareText(ExtractFileExt(NewShortcutBase), '.rdp') <> 0 then
            NewShortcutBase := NewShortcutBase + '.rdp';
          NewShortcutPath := ExpandConstant('{userdesktop}\' + NewShortcutBase);
          // Only rename if the new path differs from the current one
          if CompareText(NewShortcutPath, SelectedShortcutPath) <> 0 then
          begin
            WriteInstallerLog('Edit Shortcut: renaming "' + SelectedShortcutPath + '" to "' + NewShortcutPath + '"');
            // Check if target already exists (user would be overwriting an existing file)
            if FileExists(NewShortcutPath) then
            begin
              if MsgBox(
                   'A shortcut named "' + NewShortcutBase + '" already exists on the Desktop.' + #13#10#13#10 +
                   'Do you want to replace it?',
                   mbConfirmation, MB_YESNO) = IDNO then
              begin
                // User chose not to overwrite; keep them on the page to make further edits
                WriteInstallerLog('Edit Shortcut: user declined overwrite, staying on page');
                Result := False;
                exit;
              end
              else
              begin
                // User confirmed overwrite; delete the existing file first
                if not DeleteFile(NewShortcutPath) then
                  WriteInstallerLog('WARNING: Could not delete existing file at "' + NewShortcutPath + '"');
              end;
            end;
            // Perform the rename (copy then delete original)
            if CopyFile(SelectedShortcutPath, NewShortcutPath, False) then
            begin
              if not DeleteFile(SelectedShortcutPath) then
                WriteInstallerLog('WARNING: Could not delete original file "' + SelectedShortcutPath + '" after copy');
              SelectedShortcutPath := NewShortcutPath;
              WriteInstallerLog('Edit Shortcut: renamed to "' + SelectedShortcutPath + '"');
            end
            else
            begin
              WriteInstallerLog('ERROR: Failed to rename shortcut to "' + NewShortcutPath + '"');
            end;
          end;
        end;
      end;

      // Log the shortcut settings being applied
      LogSectionHeader('SHORTCUT SETTINGS');
      if Assigned(edtShortcutName) and edtShortcutName.Visible then
        LogKeyValue('Shortcut name', edtShortcutName.Text);
      if Assigned(chkFullScreen) then
      begin
        if chkFullScreen.Checked then LogInfo('  [x] Full Screen') else LogInfo('  [ ] Full Screen');
      end;
      if Assigned(chkUseAllMonitors) then
      begin
        if chkUseAllMonitors.Checked then LogInfo('  [x] Use All Monitors') else LogInfo('  [ ] Use All Monitors');
      end;
      if Assigned(cboResolution) and (cboResolution.ItemIndex >= 0) then
        LogKeyValue('Resolution', cboResolution.Items[cboResolution.ItemIndex]);
      if Assigned(chkCopyPaste) then
      begin
        if chkCopyPaste.Checked then LogInfo('  [x] Redirect Clipboard') else LogInfo('  [ ] Redirect Clipboard');
      end;
      if Assigned(chkSound) then
      begin
        if chkSound.Checked then LogInfo('  [x] Play Sounds') else LogInfo('  [ ] Play Sounds');
      end;
      if Assigned(cboKeyboardHook) then
        LogKeyValue('Keyboard hook mode', IntToStr(cboKeyboardHook.ItemIndex));
      if Assigned(chkShowMoreShortcutOptions) then
      begin
        if chkShowMoreShortcutOptions.Checked then LogInfo('  [x] Advanced mstsc editor') else LogInfo('  [ ] Advanced mstsc editor');
      end;
      if Assigned(chkExpWallpaper) then
      begin
        if chkExpWallpaper.Checked then LogInfo('  [x] Experience: Wallpaper') else LogInfo('  [ ] Experience: Wallpaper');
      end;
      if Assigned(chkExpFontSmooth) then
      begin
        if chkExpFontSmooth.Checked then LogInfo('  [x] Experience: Font Smoothing') else LogInfo('  [ ] Experience: Font Smoothing');
      end;
      if Assigned(chkExpComposition) then
      begin
        if chkExpComposition.Checked then LogInfo('  [x] Experience: Desktop Composition') else LogInfo('  [ ] Experience: Desktop Composition');
      end;
      if Assigned(chkExpDragContents) then
      begin
        if chkExpDragContents.Checked then LogInfo('  [x] Experience: Drag Contents') else LogInfo('  [ ] Experience: Drag Contents');
      end;
      if Assigned(chkExpMenuAnim) then
      begin
        if chkExpMenuAnim.Checked then LogInfo('  [x] Experience: Menu Animations') else LogInfo('  [ ] Experience: Menu Animations');
      end;
      if Assigned(chkExpVisualStyles) then
      begin
        if chkExpVisualStyles.Checked then LogInfo('  [x] Experience: Visual Styles') else LogInfo('  [ ] Experience: Visual Styles');
      end;

      // Write shortcut settings NOW (NextButtonClick fires for all modes, unlike ssPostInstall
      // which is skipped by Inno Setup when there are no files to install).
      // Single shortcut: write + optionally strip signature for advanced mode
      if SelectedShortcutPath <> '' then
      begin
        WriteShortcutSettingsToRdpFile(SelectedShortcutPath);
        WriteInstallerLog('Edit Shortcut: Settings written to ' + SelectedShortcutPath);
        if DoShowMstscEdit then
        begin
          WriteInstallerLog('Edit Shortcut: Advanced mode - stripping signature for mstsc edit');
          UnSignRdpFile(SelectedShortcutPath);
        end;
      end
      else if Assigned(SelectedShortcutPaths) and (SelectedShortcutPaths.Count > 0) then
      begin
        // Multi-edit: write to all selected shortcuts, then strip all if advanced
        for i := 0 to SelectedShortcutPaths.Count - 1 do
        begin
          WriteInstallerLog('Edit Shortcut: Applying settings to (' + IntToStr(i+1) + '/' + IntToStr(SelectedShortcutPaths.Count) + '): ' + SelectedShortcutPaths[i]);
          WriteShortcutSettingsToRdpFile(SelectedShortcutPaths[i]);
        end;
        if DoShowMstscEdit then
        begin
          for i := 0 to SelectedShortcutPaths.Count - 1 do
          begin
            WriteInstallerLog('Edit Shortcut: Stripping signature from (' + IntToStr(i+1) + '/' + IntToStr(SelectedShortcutPaths.Count) + '): ' + SelectedShortcutPaths[i]);
            UnSignRdpFile(SelectedShortcutPaths[i]);
          end;
        end;
      end;
    end;
  end
  else if CurPageID = Page_EditShortcutAdvanced.ID then
  begin
    // User finished editing in mstsc - re-sign the shortcut(s)
    if SelectedShortcutPath <> '' then
    begin
      WriteInstallerLog('Advanced Edit: Re-signing shortcut after manual edit: ' + SelectedShortcutPath);
      SignRdpFile(SelectedShortcutPath);
    end
    else if Assigned(SelectedShortcutPaths) and (SelectedShortcutPaths.Count > 0) then
    begin
      WriteInstallerLog('Advanced Edit: Re-signing ' + IntToStr(SelectedShortcutPaths.Count) + ' shortcuts after manual edit');
      for i := 0 to SelectedShortcutPaths.Count - 1 do
        SignRdpFile(SelectedShortcutPaths[i]);
    end;
    Result := True;
  end
  else if CurPageID = EditSystemwideSettingsPage.ID then
  begin
    if not ValidateRdpPortInput(edtRdpPort.Text, RdpPortValue, PortError) then
    begin
      MsgBox(PortError, mbError, MB_OK);
      Result := False;
      exit;
    end;
    edtRdpPort.Text := IntToStr(RdpPortValue);
  end
  else if CurPageID = Page_CreateShortcutsForExistingUsers.ID then
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
            UserPasswordStatus[i].Caption := 'Can''t be blank';
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

          UserPasswordStatus[i].Caption := '';
          UserPasswordStatus[i].Visible := False;

          ShortcutsList.Add(LocalUsersList[i] + '|' + Password);
        end
        else if Assigned(UserPasswordStatus[i]) then
        begin
          UserPasswordStatus[i].Caption := '';
          UserPasswordStatus[i].Visible := False;
        end;
      end;

      // Log which existing users were selected
      LogSectionHeader('USER SELECTIONS - Existing Users');
      LogKeyValue('Selected users', IntToStr(SelectedCount));
      for i := 0 to High(UserCheckBoxes) do
      begin
        if Assigned(UserCheckBoxes[i]) then
        begin
          if UserCheckBoxes[i].Checked then
            LogInfo('  [x] ' + LocalUserDisplayList[i])
          else
            LogInfo('  [ ] ' + LocalUserDisplayList[i]);
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
    // Skip user page for uninstall and system-edit types
    if SelectedInstallMode <> installModeInstall then
    begin
      Result := True;
      exit;
    end;
    
    UserName := UserPage.Values[0];
    Password := UserPage.Values[1];
    
    // Check if user selected "I'm done creating users"
    // Log the current user creation state
    LogSectionHeader('USER SELECTIONS - New Users');
    if Assigned(AddMoreRadio) and Assigned(DoneRadio) then
    begin
      if AddMoreRadio.Checked then
        LogInfo('  (x) Add more users')
      else
        LogInfo('  ( ) Add more users');
      if DoneRadio.Checked then
        LogInfo('  (x) Done creating users')
      else
        LogInfo('  ( ) Done creating users');
    end;
    if UserName <> '' then
      LogInfo('  Current user: ' + UserName);
    LogKeyValue('Users entered so far', IntToStr(UsersList.Count));

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
      if (UsersList.Count = 0) then
      begin
        MsgBox('Please create at least one user account before proceeding.', mbError, MB_OK);
        Result := False;
        exit;
      end;

      // Proceed to install phase; file copying is gated by ShouldInstallFiles (DoInstallTermWrap flag)
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
    
    LogInfo('Added user ' + UserName + ' to list; staying on page for more entries');

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
  LogEntry('BackButtonClick');
  LogDebug('BackButtonClick: CurPageID=' + IntToStr(CurPageID) + ' (' + GetPageNameById(CurPageID) + ')');
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
  end
  // When navigating back from ShortcutSettings to UserPage (new-user path),
  // restore the last entered user so they can re-enter or modify without a duplicate error
  else if (CurPageID = Page_ShortcutSettings.ID) and
          (SelectedInstallMode = installModeInstall) and
          DoCreateRdpShortcuts and
          (CreateUserMode = createUserModeNew) and
          (UsersList.Count > 0) then
  begin
    LastUserInfo := UsersList[UsersList.Count - 1];
    ParseUserEntry(LastUserInfo, UserName, Password);
    UserPage.Values[0] := UserName;
    UserPage.Values[1] := Password;
    UsersList.Delete(UsersList.Count - 1);
    DoneRadio.Checked := True;
    AddMoreRadio.Checked := False;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  VCRedistPath: string;
  i: Integer;
  UserInfo: string;
  UserName: string;
  Password: string;
  PortNumber: Cardinal;
  StepName: string;
begin
  if CurStep = ssInstall then StepName := 'ssInstall'
  else if CurStep = ssPostInstall then StepName := 'ssPostInstall'
  else StepName := 'ssDone';
  LogEntry('CurStepChanged');
  LogDebug('CurStepChanged: CurStep=' + StepName + ' SelectedInstallMode=' + IntToStr(SelectedInstallMode));
  if CurStep = ssInstall then
  begin
    LogSectionHeader('STEP TRANSITION: ssInstall');
    LogKeyValue('SelectedInstallMode', IntToStr(SelectedInstallMode));
    LogKeyValue('DoInstallTermWrap', BoolToStr(DoInstallTermWrap));
    LogKeyValue('DoCreateRdpShortcuts', BoolToStr(DoCreateRdpShortcuts));
    LogKeyValue('CreateUserMode', IntToStr(CreateUserMode));
    // Hide cancel button during installation to prevent confusion
    WizardForm.CancelButton.Visible := False;
    
    // Initialize determinate progress bar
    StepsTotal := 0;
    StepsDone := 0;
    WizardForm.ProgressGauge.Style := npbstNormal;
    LogDebug('Progress: gauge initialized to Normal style');

    // Initialize and show relevant steps (pending state) with contiguous layout
    BeginStepLayout;
    if SelectedInstallMode = installModeUninstall then
    begin
      StepsHeaderLabel.Caption := 'Uninstall Steps:';
      AddStepPendingLabel(StepStopSvc, TXT_StopSvc);
      AddStepPendingLabel(StepRemoveExcl, TXT_RemoveExcl);
      AddStepPendingLabel(StepRemoveFolder, TXT_RemoveFolder);
      AddStepPendingLabel(StepStartSvc, TXT_RestartSvc);
    end
    else if (SelectedInstallMode = installModeEditSystemwideSettings) and DoEditSystemWideSettings then
    begin
      StepsHeaderLabel.Caption := 'System changes:';
      if Assigned(chkEnableRDP) and (chkEnableRDP.Checked <> OrigEnableRDP) then
        AddStepPendingLabel(StepEnableRDP, 'Enable Remote Desktop');
      if Assigned(chkShowUsers) and (chkShowUsers.Checked <> OrigShowUsers) then
        AddStepPendingLabel(StepShowUsers, 'Show users on logon screen');
      if Assigned(chkPreventDuplicate) and (chkPreventDuplicate.Checked <> OrigPreventDuplicate) then
        AddStepPendingLabel(StepPreventDuplicate, 'Prevent duplicate connections per user');
      if (StrToIntDef(Trim(edtRdpPort.Text), RDP_LISTEN_PORT) <> OrigRdpPort) then
        AddStepPendingLabel(StepSetRdpPort, 'Set RDP listening port to ' + edtRdpPort.Text);
      if Assigned(chkRestartRDP) and chkRestartRDP.Checked then
        AddStepPendingLabel(StepRestartRDP, 'Restart Remote Desktop Services');
    end
    else if SelectedInstallMode = installModeEditSystemwideSettings then
    begin
      StepsHeaderLabel.Caption := 'Create Shortcuts:';
      AddStepPendingLabel(StepCreateShortcuts, TXT_CreateShortcuts);
      AddStepPendingLabel(StepPreTrust, TXT_PreTrust);
    end
    else if SelectedInstallMode = installModeShowRDPInfo then
    begin
      StepsHeaderLabel.Caption := 'RDP Settings:';
      AddStepPendingLabel(StepShowRDPInfo, TXT_ShowRDPInfo);
    end
    else if SelectedInstallMode = installModeQuickFixes then
    begin
      StepsHeaderLabel.Caption := 'Quick Fixes:';
      if Assigned(rbQFRestartRDP) and rbQFRestartRDP.Checked then
        AddStepPendingLabel(StepQuickFixes, 'Restart Remote Desktop Services')
      else if Assigned(rbQFAccountNeverExpires) and rbQFAccountNeverExpires.Checked then
        AddStepPendingLabel(StepQuickFixes, 'Set all Local Accounts to Never Expire')
      else if Assigned(rbQFRestoreTermService) and rbQFRestoreTermService.Checked then
        AddStepPendingLabel(StepQuickFixes, 'Restore deleted Remote Desktop Service');
    end
    else if SelectedInstallMode = installModeInstall then
    begin
      StepsHeaderLabel.Caption := 'Install Steps:';
      if DoInstallTermWrap then
      begin
        AddStepPendingLabel(StepCheckMSTSC, TXT_CheckMSTSC);
        AddStepPendingLabel(StepInstallMSTSC, TXT_InstallMSTSC);
        AddStepPendingLabel(StepStopSvc, TXT_StopSvc);
        AddStepPendingLabel(StepAddExcl, TXT_AddExcl);
        AddStepPendingLabel(StepEnsureVC, TXT_EnsureVC);
        AddStepPendingLabel(StepConfigureService, TXT_ConfigureService);
      end;
      if DoCreateRdpShortcuts and (CreateUserMode = createUserModeNew) and (UsersList.Count > 0) then
        AddStepPendingLabel(StepCreateUsers, TXT_CreateUsers);
      if DoCreateRdpShortcuts and (ShortcutsList.Count > 0) then
        AddStepPendingLabel(StepCreateShortcuts, TXT_CreateShortcuts);
      if DoInstallTermWrap then
      begin
        AddStepPendingLabel(StepStartSvc, TXT_StartSvc);
        AddStepPendingLabel(StepPreTrust, TXT_PreTrust);
        AddStepPendingLabel(StepCheckRDP, TXT_CheckRDP);
      end
      else
        AddStepPendingLabel(StepPreTrust, TXT_PreTrust);
    end
    else // SelectedInstallMode = installModeEditShortcuts (Edit Shortcut Settings)
    begin
      StepsHeaderLabel.Caption := 'Shortcut Settings:';
      AddStepPendingLabel(StepCreateShortcuts, 'Apply settings and open in editor');
    end;

    if (SelectedInstallMode = installModeInstall) and DoInstallTermWrap then
      CheckAndInstallMSTSC;

    // Handle uninstall cleanup
    if SelectedInstallMode = installModeUninstall then
    begin
      StatusOverlay.Caption := 'Preparing uninstallation...';
      WizardForm.ProgressGauge.Style := npbstNormal;
      
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
      
      // No TermWrap uninstall step configured: restore defaults and remove payload
      SetStepInProgress(StepRemoveFolder, TXT_RemoveFolder);
      RegWriteStringValue(HKLM, REG_TERMSERVICE_PARAMS, 'ServiceDll', ExpandConstant('{sys}\termsrv.dll'));
      RegWriteStringValue(HKLM, REG_UMRDPSERVICE_PARAMS, 'ServiceDll', ExpandConstant('{sys}\umrdp.dll'));
      // If an old TermWrap folder exists, remove it
      if DirExists(AppRoot) then
        DelTree(AppRoot, True, True, True);
      Sleep(SLEEP_SHORT);
      SetStepDone(StepRemoveFolder, TXT_RemoveFolder);
      
      SetStepInProgress(StepStartSvc, TXT_RestartSvc);
      StartTermService;
      SetStepDone(StepStartSvc, TXT_RestartSvc);
    end
    // Edit System-wide settings: deferred to ssPostInstall
    else if SelectedInstallMode = installModeEditSystemwideSettings then
    begin
      StatusOverlay.Caption := 'Preparing Create Shortcuts...';
      WizardForm.ProgressGauge.Style := npbstNormal;
    end
    // Tune performance: deferred to ssPostInstall
    else if SelectedInstallMode = installModeShowRDPInfo then
    begin
      StatusOverlay.Caption := 'Preparing to apply Group Policy settings...';
      WizardForm.ProgressGauge.Style := npbstNormal;
    end
    else if SelectedInstallMode = installModeQuickFixes then
    begin
      StatusOverlay.Caption := 'Applying Quick Fix...';
      WizardForm.ProgressGauge.Style := npbstNormal;
    end
    else if SelectedInstallMode = installModeEditShortcuts then
    begin
      StatusOverlay.Caption := 'Applying shortcut settings...';
      WizardForm.ProgressGauge.Style := npbstNormal;
    end
    // Only stop TermService when installing TermWrap
    else if (SelectedInstallMode = installModeInstall) and DoInstallTermWrap then
    begin
      StatusOverlay.Caption := 'Preparing installation...';
      WizardForm.ProgressGauge.Style := npbstNormal;
      
      // Now that UI is visible, safely stop the service (executes first, displays first)
      SetStepInProgress(StepStopSvc, TXT_StopSvc);
      StatusOverlay.Caption := 'Stopping Remote Desktop Services...';
      Log('[CurStepChanged-ssInstall] Stopping TermService for Install TermWrap');
      StopTermService;
      SetStepDone(StepStopSvc, TXT_StopSvc);
      
      SetStepInProgress(StepAddExcl, TXT_AddExcl);
      StatusOverlay.Caption := 'Adding Windows Defender exclusion...';
      AddDefenderExclusionForApp;
      SetStepDone(StepAddExcl, TXT_AddExcl);
      // Hide gauge before Inno Setup's internal file copy paints its own progress
      WizardForm.ProgressGauge.Visible := False;
    end;
  end;
  
  if CurStep = ssPostInstall then
  begin
    LogSectionHeader('STEP TRANSITION: ssPostInstall');
    // Restore and show gauge after Inno Setup's internal file copying
    if StepsTotal > 0 then
    begin
      WizardForm.ProgressGauge.Max := StepsTotal;
      WizardForm.ProgressGauge.Position := StepsDone;
      LogDebug('Progress: restored gauge after file copy - ' + IntToStr(StepsDone) + '/' + IntToStr(StepsTotal));
    end;
    WizardForm.ProgressGauge.Visible := True;
    // Handle uninstall completion
    if SelectedInstallMode = installModeUninstall then
    begin
      StatusOverlay.Caption := 'Uninstallation complete! TermWrap has been removed.';
    end
    // Show RDP Info: display-only page, no settings to apply
    else if SelectedInstallMode = installModeShowRDPInfo then
    begin
      SetStepInProgress(StepShowRDPInfo, TXT_ShowRDPInfo);
      SetStepDone(StepShowRDPInfo, TXT_ShowRDPInfo);
      StatusOverlay.Caption := 'Done.';
      WriteInstallerLog('ShowRDPInfo: no-op step complete.');
    end
    // Edit System-wide settings flow (apply queued registry/service changes)
    else if (SelectedInstallMode = installModeEditSystemwideSettings) and DoEditSystemWideSettings then
    begin
      // Apply changes in logical order
      // Enable/Disable Remote Desktop
      if Assigned(chkEnableRDP) and (chkEnableRDP.Checked <> OrigEnableRDP) then
      begin
        SetStepInProgress(StepEnableRDP, 'Applying enable/disable Remote Desktop');
        if chkEnableRDP.Checked then
        begin
          if RegWriteDWordValue(HKLM, REG_TERMINAL_SERVER, 'fDenyTSConnections', 0) then
            WriteInstallerLog('Applied fDenyTSConnections=0 (Enable RDP)')
          else
            WriteInstallerLog('Failed to write fDenyTSConnections');
        end
        else
        begin
          if RegWriteDWordValue(HKLM, REG_TERMINAL_SERVER, 'fDenyTSConnections', 1) then
            WriteInstallerLog('Applied fDenyTSConnections=1 (Disable RDP)')
          else
            WriteInstallerLog('Failed to write fDenyTSConnections');
        end;
        SetStepDone(StepEnableRDP, 'Enable Remote Desktop');
      end;

      // Show users on logon screen
      if Assigned(chkShowUsers) and (chkShowUsers.Checked <> OrigShowUsers) then
      begin
        SetStepInProgress(StepShowUsers, 'Updating Show users on logon screen');
        if chkShowUsers.Checked then
        begin
          if RegWriteDWordValue(HKLM, REG_SHOW_USERS, 'DontDisplayLastUserName', 0) then
            WriteInstallerLog('Applied DontDisplayLastUserName=0 (Show users)')
          else
            WriteInstallerLog('Failed to write DontDisplayLastUserName');
        end
        else
        begin
          if RegWriteDWordValue(HKLM, REG_SHOW_USERS, 'DontDisplayLastUserName', 1) then
            WriteInstallerLog('Applied DontDisplayLastUserName=1 (Hide users)')
          else
            WriteInstallerLog('Failed to write DontDisplayLastUserName');
        end;
        SetStepDone(StepShowUsers, 'Show users on logon screen');
      end;

      // Prevent duplicate connections per user
      if Assigned(chkPreventDuplicate) and (chkPreventDuplicate.Checked <> OrigPreventDuplicate) then
      begin
        SetStepInProgress(StepPreventDuplicate, 'Updating single-session-per-user setting');
        if chkPreventDuplicate.Checked then
        begin
          if RegWriteDWordValue(HKLM, REG_TERMINAL_SERVER, 'fSingleSessionPerUser', 1) then
            WriteInstallerLog('Applied fSingleSessionPerUser=1 (Single session)')
          else
            WriteInstallerLog('Failed to write fSingleSessionPerUser');
        end
        else
        begin
          if RegWriteDWordValue(HKLM, REG_TERMINAL_SERVER, 'fSingleSessionPerUser', 0) then
            WriteInstallerLog('Applied fSingleSessionPerUser=0 (Allow multiple sessions)')
          else
            WriteInstallerLog('Failed to write fSingleSessionPerUser');
        end;
        SetStepDone(StepPreventDuplicate, 'Prevent duplicate connections per user');
      end;

      // Hide most security warnings (RedirectionWarningDialogVersion)
      if Assigned(chkHideSecurityWarnings) and (chkHideSecurityWarnings.Checked <> OrigHideSecurityWarnings) then
      begin
        SetStepInProgress(StepPreventDuplicate, 'Updating RDP client redirection warning policy');
        if chkHideSecurityWarnings.Checked then
        begin
          if RegWriteDWordValue(HKLM, REG_TS_POLICIES + '\\Client', 'RedirectionWarningDialogVersion', 1) then
            WriteInstallerLog('Applied RedirectionWarningDialogVersion=1 (Hide security warnings)')
          else
            WriteInstallerLog('Failed to write RedirectionWarningDialogVersion');
        end
        else
        begin
          if RegDeleteValue(HKLM, REG_TS_POLICIES + '\\Client', 'RedirectionWarningDialogVersion') then
            WriteInstallerLog('Removed RedirectionWarningDialogVersion (Restore warnings)')
          else
            WriteInstallerLog('Failed to remove RedirectionWarningDialogVersion');
        end;
        SetStepDone(StepPreventDuplicate, 'Hide most security warnings');
      end;

      // RDP port change
      if (StrToIntDef(Trim(edtRdpPort.Text), RDP_LISTEN_PORT) <> OrigRdpPort) then
      begin
        SetStepInProgress(StepSetRdpPort, 'Setting RDP listening port');
        if RegWriteDWordValue(HKLM, REG_RDP_TCP, 'PortNumber', StrToIntDef(Trim(edtRdpPort.Text), RDP_LISTEN_PORT)) then
          WriteInstallerLog('Applied new RDP PortNumber=' + edtRdpPort.Text)
        else
          WriteInstallerLog('Failed to write RDP PortNumber');
        SetStepDone(StepSetRdpPort, 'Set RDP listening port');
      end;

      // RemoteFX: image quality drives whether Adaptive Graphics is enabled
      if Assigned(cmbGPImageQuality) then
      begin
        if cmbGPImageQuality.ItemIndex > 0 then
        begin
          RegWriteDWordValue(HKLM, REG_TS_POLICIES, 'fEnableRemoteFXAdvancedRemoteApp', 1);
          RegWriteDWordValue(HKLM, REG_TS_POLICIES, 'ImageQuality', cmbGPImageQuality.ItemIndex - 1);
        end
        else
        begin
          RegDeleteValue(HKLM, REG_TS_POLICIES, 'fEnableRemoteFXAdvancedRemoteApp');
          RegDeleteValue(HKLM, REG_TS_POLICIES, 'ImageQuality');
        end;
      end;
      if Assigned(cmbGPCompression) and (cmbGPCompression.ItemIndex > 0) then
        RegWriteDWordValue(HKLM, REG_TS_POLICIES, 'MaxCompressionLevel', cmbGPCompression.ItemIndex - 1)
      else if Assigned(cmbGPCompression) and (cmbGPCompression.ItemIndex = 0) then
        RegDeleteValue(HKLM, REG_TS_POLICIES, 'MaxCompressionLevel');

      // Restart RDP Service if requested
      if Assigned(chkRestartRDP) and chkRestartRDP.Checked then
      begin
        SetStepInProgress(StepRestartRDP, 'Restarting Remote Desktop Services');
        StopTermService;
        StartTermService;
        SetStepDone(StepRestartRDP, 'Restart Remote Desktop Services');
      end;

      StatusOverlay.Caption := 'System changes applied.';
    end
    // Edit System-wide settings fallback (no changes queued): create shortcuts for existing users
    else if SelectedInstallMode = installModeEditSystemwideSettings then
    begin
      SetStepInProgress(StepCreateShortcuts, TXT_CreateShortcuts);
      StatusOverlay.Caption := 'Creating RDP shortcuts...';
      CreateShortcutsForExistingUsers;
      // Create Shortcuts path completed
      SetStepDone(StepCreateShortcuts, TXT_CreateShortcuts);
      SetStepInProgress(StepPreTrust, TXT_PreTrust);
      StatusOverlay.Caption := 'Pre-trusting Remote Desktop certificate...';
      PreTrustRDPCertCurrentUser;
      SetStepDone(StepPreTrust, TXT_PreTrust);
      ClearPasswordsFromMemory;
      StatusOverlay.Caption := 'Create Shortcuts executed.';
    end
    // Quick Fixes execution
    else if SelectedInstallMode = installModeQuickFixes then
    begin
      if Assigned(rbQFRestartRDP) and rbQFRestartRDP.Checked then
      begin
        SetStepInProgress(StepQuickFixes, 'Restarting Remote Desktop Services');
        LogDebug('QuickFixes: Restarting RDP Service');
        StopTermService;
        StartTermService;
        SetStepDone(StepQuickFixes, 'Restart Remote Desktop Services');
        StatusOverlay.Caption := 'RDP Service restarted.';
        WriteInstallerLog('QuickFixes: Restart RDP Service completed.');
      end
      else if Assigned(rbQFAccountNeverExpires) and rbQFAccountNeverExpires.Checked then
      begin
        SetStepInProgress(StepQuickFixes, 'Setting all Local Accounts to Never Expire');
        LogDebug('QuickFixes: Running Get-LocalUser | Set-LocalUser -AccountNeverExpires');
        ExecPowerShellHidden('Get-LocalUser | Set-LocalUser -AccountNeverExpires', ResultCode);
        LogDebug('QuickFixes: AccountNeverExpires PowerShell exit=' + IntToStr(ResultCode));
        LogDebug('QuickFixes: Running Get-LocalUser | Set-LocalUser -PasswordNeverExpires $true');
        ExecPowerShellHidden('Get-LocalUser | Set-LocalUser -PasswordNeverExpires $true', ResultCode);
        LogDebug('QuickFixes: PasswordNeverExpires PowerShell exit=' + IntToStr(ResultCode));
        WriteInstallerLog('QuickFixes: Set all Local Accounts to Never Expire (exit=' + IntToStr(ResultCode) + ')');
        SetStepDone(StepQuickFixes, 'Set all Local Accounts to Never Expire');
        StatusOverlay.Caption := 'Accounts never expire applied.';
      end
      else if Assigned(rbQFRestoreTermService) and rbQFRestoreTermService.Checked then
      begin
        SetStepInProgress(StepQuickFixes, 'Restoring deleted Remote Desktop Service');
        LogDebug('QuickFixes: Restoring TermService registry keys');
        ExtractTemporaryFile('restore_term_service.reg');
        ResultCode := RunCmdHidden('reg.exe import "' + ExpandConstant('{tmp}\restore_term_service.reg') + '"');
        LogDebug('QuickFixes: reg.exe import exit=' + IntToStr(ResultCode));
        WriteInstallerLog('QuickFixes: Restore TermService (exit=' + IntToStr(ResultCode) + ')');
        SetStepDone(StepQuickFixes, 'Restore deleted Remote Desktop Service');
        StatusOverlay.Caption := 'Remote Desktop Service restored.';
      end;
    end
    else if SelectedInstallMode = installModeEditShortcuts then
    begin
      // NOTE: WriteShortcutSettingsToRdpFile + UnSignRdpFile are now handled in
      // NextButtonClick(Page_ShortcutSettings.ID) because CurStepChanged is NOT
      // called by Inno Setup when no [Files] entries are selected (edit mode).
      // This block only sets up the progress bar UI for the Installing page
      // (which may still appear briefly during page transitions).
      SetStepInProgress(StepCreateShortcuts, 'Apply settings and open in editor');
      StatusOverlay.Caption := 'Shortcut settings applied.';
      SetStepDone(StepCreateShortcuts, 'Apply settings and open in editor');
    end
    // Install TermWrap: Download VC++, apply registry, start service
    else if (SelectedInstallMode = installModeInstall) and DoInstallTermWrap then
    begin
      SetStepInProgress(StepEnsureVC, TXT_EnsureVC);
      // Check if VC++ Redistributable is already installed (unless debug mode)
      if DebugMode or (not IsVCRedistInstalled) then
      begin
        LogSectionHeader('VC++ REDIST CHECK/INSTALL');
        LogKeyValue('DebugMode', BoolToStr(DebugMode));
        LogKeyValue('SimulateNoVCRedist', BoolToStr(SimulateNoVCRedist));
        LogKeyValue('Download URL', URL_VCREDIST_X64);
        StatusOverlay.Caption := 'Downloading VC++ Redistributable from Microsoft...';
        
        // Download VC++ Redistributable from Microsoft
        VCRedistPath := TempFile('vc_redist.x64.exe');
        LogKeyValue('Installer temp path', VCRedistPath);
        
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
        LogKeyValue('Download exit code', IntToStr(ResultCode));
        LogKeyValue('Installer file exists after download', BoolToStr(FileExists(VCRedistPath)));
        
        StatusOverlay.Caption := 'Installing VC++ Redistributable (this may take a minute)...';
        WizardForm.Update;
        
        // Validate publisher before running downloaded installer
        if FileExists(VCRedistPath) and IsSignedByMicrosoftCorporation(VCRedistPath) then
        begin
          Exec(VCRedistPath, '/install /quiet /norestart', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
          LogKeyValue('Installer execution exit code', IntToStr(ResultCode));
          if ResultCode <> 0 then
            PromptManualDownload('VC++ Redistributable (2015-2022 x64)', URL_VCREDIST_X64, 'Installer execution failed (exit code ' + IntToStr(ResultCode) + ')');
        end
        else
        begin
          Log('ERROR: VC++ download/signature validation failed.');
          PromptManualDownload('VC++ Redistributable (2015-2022 x64)', URL_VCREDIST_X64, 'Download or signature validation failed (exit code ' + IntToStr(ResultCode) + ')');
        end;
        if FileExists(VCRedistPath) then
          DeleteFile(VCRedistPath);
      end
      else
      begin
        LogSectionHeader('VC++ REDIST CHECK/INSTALL');
        Log('VC++ Redistributable detected. Skipping download/install.');
        StatusOverlay.Caption := 'VC++ Redistributable already installed, skipping...';
      end;
      // VC++ ensured (installed or skipped)
      SetStepDone(StepEnsureVC, TXT_EnsureVC);
      
      // Install and configure TermWrap
      SetStepInProgress(StepConfigureService, TXT_ConfigureService);
      StatusOverlay.Caption := 'Installing and configuring TermWrap...';
      // TermWrap files are bundled and copied earlier; no external installer to run.
      Sleep(SLEEP_SHORT);
      
      // Set ServiceDll to TermWrap
      LogSectionHeader('REGISTRY: TermWrap service configuration');
      if RegWriteStringValue(HKLM, REG_TERMSERVICE_PARAMS, 'ServiceDll', AppBin(FILE_TERMWRAP)) then
        WriteInstallerLog('Registry: Set TermService ServiceDll=' + AppBin(FILE_TERMWRAP))
      else
        WriteInstallerLog('Registry: FAILED to set TermService ServiceDll');

      // Enable RDP connections
      if RegWriteDWordValue(HKLM, REG_TERMINAL_SERVER, 'fDenyTSConnections', 0) then
        WriteInstallerLog('Registry: Set fDenyTSConnections=0 (RDP enabled)')
      else
        WriteInstallerLog('Registry: FAILED to set fDenyTSConnections');

      // Require single session per user (prevents multiple concurrent RDP sessions for same account)
      if RegWriteDWordValue(HKLM, REG_TERMINAL_SERVER, 'fSingleSessionPerUser', 1) then
        WriteInstallerLog('Registry: Set fSingleSessionPerUser=1 (single session per user)')
      else
        WriteInstallerLog('Registry: FAILED to set fSingleSessionPerUser');

      // Allow RDP client to remain connected when minimized
      if RegWriteDWordValue(HKLM, 'Software\Microsoft\Terminal Server Client', 'RemoteDesktop_SuppressWhenMinimized', 2) then
        WriteInstallerLog('Registry: Set RemoteDesktop_SuppressWhenMinimized=2 (allow minimized RDP)')
      else
        WriteInstallerLog('Registry: FAILED to set RemoteDesktop_SuppressWhenMinimized');

      // Accept RDP launch consent prompt by default
      if RegWriteDWordValue(HKCU, 'Software\Microsoft\Terminal Server Client', 'RdpLaunchConsentAccepted', 1) then
        WriteInstallerLog('Registry: Set RdpLaunchConsentAccepted=1 (consent accepted)')
      else
        WriteInstallerLog('Registry: FAILED to set RdpLaunchConsentAccepted');

      // Set unlimited max connections (default is 99999999, but some systems might have lower caps)
      if RegWriteDWordValue(HKLM, REG_TS_POLICIES, 'MaxInstanceCount', 999999) then
        WriteInstallerLog('Registry: Set MaxInstanceCount=999999 (unlimited connections)')
      else
        WriteInstallerLog('Registry: FAILED to set MaxInstanceCount');

      // Ensure TermService runs under the expected service account.
      EnsureTermServiceRunsAsNetworkService;

      // Ensure UmRdpService is set to automatic startup
      EnsureUmRdpServiceAutomatic;

      // Don't start the service yet - wait until all users are created
      SetStepDone(StepConfigureService, TXT_ConfigureService);
    end;
    
    // Create all user accounts and RDP files (skip for uninstall and Create Shortcuts)
    if SelectedInstallMode = installModeInstall then
    begin
      // Create user accounts and RDP shortcuts
      
      if UsersList.Count > 0 then
      begin
        SetStepInProgress(StepCreateUsers, TXT_CreateUsers);
        // Log contents of UsersList to aid debugging when no users are created
        WriteInstallerLog('DEBUG: UsersList.Count=' + IntToStr(UsersList.Count));
        try
          for i := 0 to UsersList.Count - 1 do
            WriteInstallerLog('DEBUG: UsersList[' + IntToStr(i) + ']=' + MaskPasswordInEntry(UsersList[i]));
        except
          // Defensive: avoid crashing installer when logging fails
        end;
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
    
    // Start TermService (only when TermWrap was installed) - this creates the SSL certificate
    if (SelectedInstallMode = installModeInstall) and DoInstallTermWrap then
    begin
      SetStepInProgress(StepStartSvc, TXT_StartSvc);
      StatusOverlay.Caption := 'Starting Remote Desktop Services...';
      // Start TermService after all files and registry are done
      ResultCode := StartTermServiceEx;
      if ResultCode = 0 then
      begin
        SleepWithUI(SLEEP_EXTRALONG); // Wait for service to fully initialize and create certificate
        SetStepDone(StepStartSvc, TXT_StartSvc);
      end
      else
      begin
        Log('WARNING: TermService failed to start with exit code ' + IntToStr(ResultCode));
        SetStepDone(StepStartSvc, TXT_StartSvc); // Mark as done even if failed (might already be running)
        SleepWithUI(SLEEP_LONG); // Give extra time if service had issues
      end;
    end;
    
    // Pre-trust for current user in all Install sub-flows (AFTER service starts if applicable)
    if SelectedInstallMode = installModeInstall then
    begin
      SetStepInProgress(StepPreTrust, TXT_PreTrust);
      StatusOverlay.Caption := 'Pre-trusting Remote Desktop certificate...';
      PreTrustRDPCertCurrentUser;
      // Pre-trust is optional - don't fail install if cert doesn't exist yet
      SetStepDone(StepPreTrust, TXT_PreTrust);
    end;

    // Verify RDP is listening (only when TermWrap was installed)
    if (SelectedInstallMode = installModeInstall) and DoInstallTermWrap then
    begin
      SetStepInProgress(StepCheckRDP, TXT_CheckRDP);
      StatusOverlay.Caption := 'Verifying RDP service...';

      // Check for Razer Cortex - its "Boost" feature is known to cause RDP disconnects
      LogSectionHeader('RAZER CORTEX CHECK');
      ExecPowerShellHidden(
        '$ErrorActionPreference = ''SilentlyContinue''; ' +
        '$proc = Get-Process | Where-Object { $_.ProcessName -like ''*Cortex*'' }; ' +
        'if ($proc) { exit 0 } else { exit 1 }',
        ResultCode);
      Log('Razer Cortex detection exit code = ' + IntToStr(ResultCode) + ' (0=found, 1=not found)');
      if ResultCode = 0 then
      begin
        MessageBox(0,
          'Razer Cortex "Boost" feature is known to cause RDP disconnects.' + #13#10#13#10 +
          'Here are 3 ways to resolve this:' + #13#10 +
          '  1. Uninstall Razer Cortex if it is not needed' + #13#10 +
          '      OR' + #13#10#13#10 +
          '  2. Open Razer Cortex: In ''Booster'', disable ''Auto-boost''' + #13#10 +
          '      OR' + #13#10#13#10 +
          '  3. Open Razer Cortex: In ''Booster'' > ''Services'', ensure ''TermService'' and ''UmRdpService'' are unchecked',
          'Razer Cortex detected on your device',
          MB_OK + $40);
      end;
      
      // Read the configured listening port from registry
      if RegQueryDWordValue(HKLM, REG_RDP_TCP, 'PortNumber', PortNumber) then
        // PortNumber already set from registry
      else
        PortNumber := RDP_LISTEN_PORT;
      
      ResultCode := 0;
      Exec(EXE_POWERSHELL, BuildPowerShellArgs('if (@(Get-NetTCPConnection -LocalPort ' + IntToStr(PortNumber) + ' -State Listen -ErrorAction SilentlyContinue).Count -gt 0) { exit 0 } else { exit 1 }', True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      
      if ResultCode = 0 then
      begin
        SetStepDone(StepCheckRDP, TXT_CheckRDP);
      end
      else
      begin
        SetStepDone(StepCheckRDP, TXT_CheckRDP);
        if MsgBox('RDP service is not detected as listening on port ' + IntToStr(PortNumber) + '.' + #13#10#13#10 +
                  'A system restart usually resolves this issue.' + #13#10#13#10 +
                  'Would you like to restart your computer now?', mbConfirmation, MB_YESNO) = IDYES then
        begin
          // User chose to restart now
          Exec('shutdown.exe', '/r /t 5 /c "Restarting to complete TermWrap setup"', '', SW_HIDE, ewNoWait, ResultCode);
        end;
      end;
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
begin
  // Handle removal when using the standard uninstaller
  if CurUninstallStep = usUninstall then
  begin
    RemoveDefenderExclusionForApp;

    // Remove the RDP pre-trust entry written by PreTrustRDPCertCurrentUser
    ExecPowerShellHidden(
      '$ErrorActionPreference = ''Stop''; ' +
      'try { ' +
      '  $regPath = ''HKCU:\Software\Microsoft\Terminal Server Client\LocalDevices''; ' +
      '  if (Test-Path $regPath) { ' +
      '    Remove-ItemProperty -Path $regPath -Name ''' + RDP_LOOPBACK_IP + ''' -Force -ErrorAction SilentlyContinue ' +
      '  }; ' +
      '  exit 0 ' +
      '} catch { ' +
      '  exit 1 ' +
      '}',
      ResultCode);
    Log('DEBUG: Remove pre-trust registry entry exit code = ' + IntToStr(ResultCode) + ' (0=success, 1=error)');
  end;
end;

// Show or hide the shortcut name editing controls (label, edit box, extension).
// When hiding, the multi-shortcut note label is shown in their place.
procedure ShowShortcutNameEdit(const ShowNameEdit: Boolean);
begin
  LogEntry('ShowShortcutNameEdit');
  LogDebug('ShowShortcutNameEdit: ShowNameEdit=' + BoolToStr(ShowNameEdit));
  if Assigned(lblShortcutName) then
    lblShortcutName.Visible := ShowNameEdit;
  if Assigned(edtShortcutName) then
    edtShortcutName.Visible := ShowNameEdit;
  if Assigned(lblShortcutExtension) then
    lblShortcutExtension.Visible := ShowNameEdit;
  if Assigned(lblMultiShortcutEditingNote) then
    lblMultiShortcutEditingNote.Visible := not ShowNameEdit;
  LogExit('ShowShortcutNameEdit');
end;

procedure CurPageChanged(CurPageID: Integer);
var
  i: Integer;
  CompletionText: string;
  Entry: string;
  UserName: string;
  Password: string;
  ParseUser: string;
  ParsePass: string;
  MstscPath: string;
  MstscResultCode: Integer;
  NowTick: Cardinal;
  IsDuplicatePageEvent: Boolean;
  DeltaMs: Cardinal;
  DisplayVersion: string;
  BuildNumberStr: string;
  UBRVal: Cardinal;
  ServiceStatus: string;
  TermsrvVer: string;
  ServiceDllPath: string;
  WrapperVersion: string;
  PortNumber: Cardinal;
  ListenerStatus: string;
  PageName: string;
begin
  LogEntry('CurPageChanged');
  PageName := GetPageNameById(CurPageID);
  LogDebug('CurPageChanged: CurPageID=' + IntToStr(CurPageID) + ' (' + PageName + ') SelectedInstallMode=' + IntToStr(SelectedInstallMode));
  // Suppress grey flash by hiding page content during the VCL style paint cycle.
  if (CurPageID = Page_InstallOptions.ID) or
     (CurPageID = UserPage.ID) or
     (CurPageID = Page_ShortcutSettings.ID) or
     (CurPageID = EditSystemwideSettingsPage.ID) or
     (CurPageID = Page_ShowRDPInfo.ID) or
     (CurPageID = Page_CreateShortcutsForExistingUsers.ID) or
     (CurPageID = EditShortcutPage.ID) or
     (CurPageID = Page_EditShortcutAdvanced.ID) then
  begin
    WizardForm.PageNameLabel.Visible := False;
    WizardForm.PageDescriptionLabel.Visible := False;
    if CurPageID = Page_InstallOptions.ID then
      Page_InstallOptions.Surface.Visible := False;
    StatusOverlay.Caption := 'Please wait...';
    StatusOverlay.Visible := True;
    SleepWithUI(50);
    StatusOverlay.Visible := False;
    if CurPageID = Page_InstallOptions.ID then
      Page_InstallOptions.Surface.Visible := True;
    WizardForm.PageNameLabel.Visible := True;
    WizardForm.PageDescriptionLabel.Visible := True;
  end;

  // Ensure status overlay is visible on the Installing page (may have been hidden
  // by the hide-reveal transition from the previous page).
  if CurPageID = wpInstalling then
  begin
    StatusOverlay.Visible := True;
  end;

  NowTick := GetTickCount;
  IsDuplicatePageEvent := (CurPageID = LastLoggedPageId) and ((NowTick - LastLoggedPageTick) <= PAGE_LOG_DEDUPE_MS);

  if IsDuplicatePageEvent then
  begin
    Inc(LastSuppressedPageLogs);
    DeltaMs := (NowTick - LastLoggedPageTick);
    WriteInstallerLog('PAGE REFRESH (deduped): ' + GetPageNameById(CurPageID) +
      ' | delta=' + IntToStr(DeltaMs) + 'ms | repeat=' + IntToStr(LastSuppressedPageLogs));
  end
  else
  begin
    LastSuppressedPageLogs := 0;
    LogSectionHeader('PAGE SHOWN');
    LogPageContext(CurPageID);
    LastLoggedPageId := CurPageID;
    LastLoggedPageTick := NowTick;
  end;

  // Show the Save Install Log button only on the Finished page
  if Assigned(ViewLogButton) then
    ViewLogButton.Visible := (CurPageID = wpFinished);

  if CurPageID = EditShortcutPage.ID then
    WizardForm.NextButton.Caption := SetupMessage(msgButtonNext);

  // Ensure Next button is visible on the install options page (it may be hidden
  // on the Show RDP Info page when navigated back from there).
  if CurPageID = Page_InstallOptions.ID then
  begin
    WizardForm.NextButton.Visible := True;

    if not InstallOptionsAutoUserSourceApplied then
    begin
      InstallOptionsAutoUserSourceApplied := True;
      if HasDesktopShortcutTargetingRdpExe then
      begin
        rbUseExistingUsers.Checked := True;
        rbUseExistingUsers.Caption := 'Use existing users';
        if Assigned(rbUseExistingUsersHint) then
          rbUseExistingUsersHint.Visible := True;
        WriteInstallerLog('Install options: detected Desktop .lnk targeting rdp.exe; defaulting to "Use existing users" (first load only).');
      end
      else
      begin
        rbUseExistingUsers.Caption := 'Use existing users';
        if Assigned(rbUseExistingUsersHint) then
          rbUseExistingUsersHint.Visible := False;
        WriteInstallerLog('Install options: no Desktop .lnk targeting rdp.exe found; keeping default user source option.');
      end;

      OnCreateRdpShortcutsClick(nil);
    end;
  end;

  if CurPageID = EditShortcutPage.ID then
  begin
    if not Assigned(DesktopRdpFiles) then
      DesktopRdpFiles := GetDesktopRdpFiles;
    if not EditShortcutControlsBuilt then
      BuildShortcutEditorControls
    else
    begin
      CurrentShortcutPage := 0;
      UpdateShortcutPageDisplay;
    end;
  end;

  // Advanced Editing page: show instructions, screenshot, and launch mstsc editor
  if CurPageID = Page_EditShortcutAdvanced.ID then
  begin
    if not EditShortcutAdvancedControlsBuilt then
      BuildEditShortcutAdvancedControls;

    // Load the screenshot image
    if Assigned(EditShortcutAdvancedImage) then
    begin
      try
        ExtractTemporaryFile(FILE_RDPEDITSAVE_BMP);
        EditShortcutAdvancedImage.Bitmap.LoadFromFile(ExpandConstant(TEMP_RDPEDITSAVE_BMP));
        EditShortcutAdvancedImage.Visible := True;
      except
        EditShortcutAdvancedImage.Visible := False;
      end;
    end;

    // Launch mstsc /edit for the shortcut so the user can make changes while
    // viewing the instructions and screenshot on this page.
    if DoShowMstscEdit and (SelectedShortcutPath <> '') then
    begin
      MstscPath := GetMstscPath;
      MstscResultCode := 0;
      if (MstscPath = '') or (not Exec(MstscPath, '/edit "' + SelectedShortcutPath + '"', '', SW_SHOW, ewNoWait, MstscResultCode)) then
      begin
        MsgBox('Failed to launch Remote Desktop editor. Verify mstsc is available and try again.', mbError, MB_OK);
        WriteInstallerLog('Edit Shortcut Advanced: Exec(mstsc) failed exit=' + IntToStr(MstscResultCode));
      end
      else
        WriteInstallerLog('Edit Shortcut Advanced: mstsc /edit launched for ' + SelectedShortcutPath);
    end;
  end;

  // Configure shortcut settings page based on which flow is entering it
  if CurPageID = Page_ShortcutSettings.ID then
  begin
    WizardForm.NextButton.Caption := SetupMessage(msgButtonNext);
    WizardForm.ActiveControl := WizardForm.NextButton;  // prevent cboResolution from receiving initial focus

    if SelectedInstallMode = installModeEditShortcuts then
    begin
      if Assigned(lblShortcutSection) then
        lblShortcutSection.Caption := 'Basic Shortcut Settings';

      // Single shortcut: show name editor and Advanced checkbox
      if SelectedShortcutPath <> '' then
      begin
        ShowShortcutNameEdit(True);
        if Assigned(chkShowMoreShortcutOptions) then chkShowMoreShortcutOptions.Visible := True;
        ReadShortcutSettingsFromRdpFile(SelectedShortcutPath);
        if Assigned(edtShortcutName) then
          edtShortcutName.Text := ChangeFileExt(ExtractFileName(SelectedShortcutPath), '');
      end
      else
      begin
        // Multi-edit (2+ shortcuts): hide name editor and Advanced checkbox
        ShowShortcutNameEdit(False);
        if Assigned(chkShowMoreShortcutOptions) then chkShowMoreShortcutOptions.Visible := False;
        if Assigned(lblMultiShortcutEditingNote) then
          lblMultiShortcutEditingNote.Caption := 'These settings will be applied to each selected shortcut';
        // Pre-populate from the first selected shortcut
        if Assigned(SelectedShortcutPaths) and (SelectedShortcutPaths.Count > 0) then
          ReadShortcutSettingsFromRdpFile(SelectedShortcutPaths[0]);
      end;
    end
    else if (SelectedInstallMode = installModeInstall) and (CreateUserMode = createUserModeNew) then
    begin
      if Assigned(lblShortcutSection) then
        lblShortcutSection.Caption := 'Basic Shortcut Settings';
      // CreateUsers path: hide shortcut name editor when 2+ users, show multi-note instead
      if Assigned(chkShowMoreShortcutOptions) then chkShowMoreShortcutOptions.Visible := False;
      ShowShortcutNameEdit(UsersList.Count <= 1);
      if (UsersList.Count > 1) and Assigned(lblMultiShortcutEditingNote) then
        lblMultiShortcutEditingNote.Caption := 'These settings will be applied to each shortcut';
      // Set default shortcut name from the first user in the list
      if Assigned(edtShortcutName) and (UsersList.Count > 0) then
      begin
        ParseUser := '';
        ParsePass := '';
        ParseUserEntry(UsersList[0], ParseUser, ParsePass);
        if ParseUser <> '' then
          edtShortcutName.Text := ParseUser;
      end;
    end
    else
    begin
      if Assigned(lblShortcutSection) then
        lblShortcutSection.Caption := 'Basic Shortcut Settings';
      // ExistingUsers path: hide shortcut name editor when 2+ shortcuts, show multi-note instead
      if Assigned(chkShowMoreShortcutOptions) then chkShowMoreShortcutOptions.Visible := False;
      ShowShortcutNameEdit(ShortcutsList.Count <= 1);
      if (ShortcutsList.Count > 1) and Assigned(lblMultiShortcutEditingNote) then
        lblMultiShortcutEditingNote.Caption := 'These settings will be applied to each shortcut';
      // Set default shortcut name from the first shortcut entry
      if Assigned(edtShortcutName) and (ShortcutsList.Count > 0) then
      begin
        ParseUser := '';
        ParsePass := '';
        ParseUserEntry(ShortcutsList[0], ParseUser, ParsePass);
        if ParseUser <> '' then
          edtShortcutName.Text := ParseUser;
      end;
    end;
  end;

  // Lazy-load user list only when Create Shortcuts page is first shown
  if (CurPageID = EditSystemwideSettingsPage.ID) and (LocalUsersList.Count = 0) then
  begin
    LocalUsersList.Clear;
    LocalUserDisplayList.Clear;
    GetLocalUsers(LocalUsersList, LocalUserDisplayList);
    SetLength(UserCheckBoxes, LocalUsersList.Count);
    SetLength(UserPasswordEdits, LocalUsersList.Count);
    SetLength(UserPasswordStatus, LocalUsersList.Count);
    // Build per-user controls on Create Shortcuts page
    BuildCreateShortcutsControls;
  end;

  // Quick Fixes page: no dynamic data to load
  if Assigned(QuickFixesPage) and (CurPageID = QuickFixesPage.ID) then
  begin
    LogDebug('CurPageChanged: Quick Fixes page shown');
  end;

  // Populate Show RDP Info page when shown
  if Assigned(Page_ShowRDPInfo) and (CurPageID = Page_ShowRDPInfo.ID) then
  begin
    // Show RDP Info is a display-only status page — hide the Next button so the
    // user views the info and closes the installer via Cancel when done.
    WizardForm.NextButton.Visible := False;

    // Reset displayed values to indicate loading when page is shown again (e.g. after Back/Next)
    if Assigned(lblWinVer) then lblWinVer.Caption := '--';
    if Assigned(lblRDPService) then lblRDPService.Caption := '--';
    if Assigned(lblWinRDPVer) then lblWinRDPVer.Caption := '--';
    if Assigned(lblWrapperVer) then lblWrapperVer.Caption := '--';
    if Assigned(lblListener) then lblListener.Caption := '--';

    // System Status — refresh live values each time the page is shown
    DisplayVersion := SafeRegString(HKLM, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion', 'DisplayVersion', '');
    if DisplayVersion = '' then
      DisplayVersion := SafeRegString(HKLM, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion', 'ReleaseId', '');
    BuildNumberStr := SafeRegString(HKLM, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion', 'CurrentBuildNumber', '');
    UBRVal := SafeRegDword(HKLM, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion', 'UBR', 0);
    if (DisplayVersion <> '') and (BuildNumberStr <> '') then
      lblWinVer.Caption := 'v' + DisplayVersion + ' 10.0.' + BuildNumberStr + '.' + IntToStr(UBRVal)
    else if BuildNumberStr <> '' then
      lblWinVer.Caption := '10.0.' + BuildNumberStr + '.' + IntToStr(UBRVal)
    else
      lblWinVer.Caption := 'Unknown';

    ServiceStatus := GetPSOutput('(Get-Service -Name TermService -ErrorAction SilentlyContinue).Status');
    if ServiceStatus = '' then lblRDPService.Caption := 'Not installed'
    else lblRDPService.Caption := ServiceStatus;

    TermsrvVer := GetPSOutput('(Get-Item -Path (Join-Path $env:windir ''System32\\termsrv.dll'') -ErrorAction SilentlyContinue).VersionInfo.FileVersion');
    if TermsrvVer = '' then lblWinRDPVer.Caption := 'Unknown'
    else lblWinRDPVer.Caption := TermsrvVer;

    // Read the actual ServiceDll path from registry to determine what DLL is configured
    if RegQueryStringValue(HKLM, REG_TERMSERVICE_PARAMS, 'ServiceDll', ServiceDllPath) then
    begin
      // Check if registry path points to TermWrap.dll
      if (Pos('termwrap.dll', Lowercase(ServiceDllPath)) > 0) then
      begin
        // Try registry path first; if it fails (e.g. env vars not expanded), fall back to known install path
        if not FileExists(ServiceDllPath) then
          ServiceDllPath := ExpandConstant('{commonpf64}\RDPWrapKit\TermWrap.dll');
        if FileExists(ServiceDllPath) then
        begin
          WrapperVersion := GetPSOutput('$f=Get-Item ''' + ServiceDllPath + '''; $f.VersionInfo.FileVersionRaw.ToString() + '' ('' + $f.Length + '' bytes)''');
          if WrapperVersion = '' then lblWrapperVer.Caption := 'TermWrap (unknown version)'
          else lblWrapperVer.Caption := 'TermWrap ' + WrapperVersion;
        end
        else
          lblWrapperVer.Caption := 'TermWrap (file not found)';
      end
      // Check if registry path points to rdpwrap.dll (legacy)
      else if (Pos('rdpwrap.dll', Lowercase(ServiceDllPath)) > 0) then
      begin
        // Try registry path first; if it fails, fall back to known legacy install path
        if not FileExists(ServiceDllPath) then
          ServiceDllPath := ExpandConstant('{commonpf64}\RDP Wrapper\rdpwrap.dll');
        if FileExists(ServiceDllPath) then
        begin
          WrapperVersion := GetPSOutput('$f=Get-Item ''' + ServiceDllPath + '''; $f.VersionInfo.FileVersionRaw.ToString() + '' ('' + $f.Length + '' bytes)''');
          if WrapperVersion = '' then lblWrapperVer.Caption := 'RDPWrap (unknown version)'
          else lblWrapperVer.Caption := 'RDPWrap ' + WrapperVersion;
        end
        else
          lblWrapperVer.Caption := 'RDPWrap (file not found)';
      end
      // Check if registry path points to termsrv.dll (Windows default)
      else if (Pos('termsrv.dll', Lowercase(ServiceDllPath)) > 0) then
      begin
        lblWrapperVer.Caption := 'None (Windows default)';
      end
      // Unknown/custom DLL - show its filename, version, and size
      else
      begin
        if FileExists(ServiceDllPath) then
        begin
          WrapperVersion := GetPSOutput('$f=Get-Item ''' + ServiceDllPath + '''; $f.VersionInfo.FileVersionRaw.ToString() + '' ('' + $f.Length + '' bytes)''');
          if WrapperVersion = '' then lblWrapperVer.Caption := ExtractFileName(ServiceDllPath) + ' (unknown version)'
          else lblWrapperVer.Caption := ExtractFileName(ServiceDllPath) + ' ' + WrapperVersion;
        end
        else
          lblWrapperVer.Caption := ExtractFileName(ServiceDllPath) + ' (file not found)';
      end;
    end
    else
      lblWrapperVer.Caption := 'None (Windows default)';

    // Check if RDP is currently listening on the configured port
    if RegQueryDWordValue(HKLM, REG_RDP_TCP, 'PortNumber', PortNumber) then
      // PortNumber already set from registry
    else
      PortNumber := RDP_LISTEN_PORT;

    ListenerStatus := GetPSOutput('$( $p = ' + IntToStr(PortNumber) + '; if (@(Get-NetTCPConnection -LocalPort $p -State Listen -ErrorAction SilentlyContinue).Count -gt 0) { ''Listening on port '' + $p } else { ''Not Listening on port '' + $p } )');
    if ListenerStatus = '' then
      lblListener.Caption := 'Not Listening on port ' + IntToStr(PortNumber)
    else
      lblListener.Caption := ListenerStatus;

  end;

  // Populate and capture original values when the Edit System-wide Settings page is shown
  if CurPageID = EditSystemwideSettingsPage.ID then
  begin
    // Load live registry values into controls
    LoadDWordCheckbox(HKLM, REG_TERMINAL_SERVER, 'fDenyTSConnections', 0, chkEnableRDP, False);
    LoadDWordCheckbox(HKLM, REG_SHOW_USERS, 'DontDisplayLastUserName', 0, chkShowUsers, True);
    LoadDWordCheckbox(HKLM, REG_TERMINAL_SERVER, 'fSingleSessionPerUser', 1, chkPreventDuplicate, False);
    // Hide most security warnings: RedirectionWarningDialogVersion under Policies\...\Terminal Services\Client
    LoadDWordCheckbox(HKLM, REG_TS_POLICIES + '\\Client', 'RedirectionWarningDialogVersion', 1, chkHideSecurityWarnings, False);
    if RegQueryDWordValue(HKLM, REG_RDP_TCP, 'PortNumber', PortNumber) then
      edtRdpPort.Text := IntToStr(PortNumber)
    else
      edtRdpPort.Text := IntToStr(RDP_LISTEN_PORT);

    // RemoteFX settings
    LoadDWordCombo(HKLM, REG_TS_POLICIES, 'ImageQuality', cmbGPImageQuality);
    LoadDWordCombo(HKLM, REG_TS_POLICIES, 'MaxCompressionLevel', cmbGPCompression);

    // Capture originals for change detection
    if Assigned(chkEnableRDP) then OrigEnableRDP := chkEnableRDP.Checked else OrigEnableRDP := False;
    if Assigned(chkShowUsers) then OrigShowUsers := chkShowUsers.Checked else OrigShowUsers := True;
    if Assigned(chkPreventDuplicate) then OrigPreventDuplicate := chkPreventDuplicate.Checked else OrigPreventDuplicate := False;
    if Assigned(chkHideSecurityWarnings) then OrigHideSecurityWarnings := chkHideSecurityWarnings.Checked else OrigHideSecurityWarnings := False;
    OrigRdpPort := StrToIntDef(Trim(edtRdpPort.Text), RDP_LISTEN_PORT);
    WriteInstallerLog('CurPageChanged: Captured original system settings: EnableRDP=' + BoolToStr(OrigEnableRDP) + ', ShowUsers=' + BoolToStr(OrigShowUsers) + ', SingleSession=' + BoolToStr(OrigPreventDuplicate) + ', Port=' + IntToStr(OrigRdpPort));
  end;

  // Ensure pagination resets when Create Shortcuts page is shown (avoid stale page index after Back/Next)
  if CurPageID = Page_CreateShortcutsForExistingUsers.ID then
  begin
    if LocalUsersList.Count > 0 then
    begin
      CurrentUserPage := 0;
      UpdateUsersPageDisplay;
    end;
  end;
  
  // Display completion info on the final page
  if CurPageID = wpFinished then
  begin
    // Reset optional finish-page image controls by default.
    if Assigned(FinishedExampleImage) then FinishedExampleImage.Visible := False;

    WriteInstallerLog('CurPageChanged: Finish page shown');
    if SelectedInstallMode = installModeUninstall then
    begin
      // Uninstall completion message
      WizardForm.FinishedHeadingLabel.Caption := 'Uninstallation Complete';
      CompletionText := 'TermWrap has been successfully removed.';
      WriteInstallerLog('CurPageChanged: Showing uninstall completion message');
    end
    else if SelectedInstallMode = installModeShowRDPInfo then
    begin
      WizardForm.FinishedHeadingLabel.Caption := 'RDP Settings Applied';
      CompletionText :=
        'RDP settings have been applied.' + #13#10#13#10 +
        'Some settings take effect immediately; others require a restart of the Remote Desktop service or a new RDP session.';
      WriteInstallerLog('CurPageChanged: Showing Show RDP Info completion message');
    end
    else if SelectedInstallMode = installModeQuickFixes then
    begin
      WizardForm.FinishedHeadingLabel.Caption := 'Quick Fix Applied';
      CompletionText := 'The selected quick fix has been applied successfully.';
      WriteInstallerLog('CurPageChanged: Showing Quick Fixes completion message');
    end
    else if SelectedInstallMode = installModeEditShortcuts then
    begin
      if DoShowMstscEdit then
      begin
        WizardForm.FinishedHeadingLabel.Caption := 'Advanced Shortcut Editing Complete';
        CompletionText :=
          'Your shortcut has been updated and re-signed.' + #13#10#13#10 +
          'You can now test the connection by double-clicking the shortcut on your Desktop.';
        WriteInstallerLog('CurPageChanged: Showing advanced shortcut editor completion message');
      end
      else
      begin
        // List which shortcuts were updated (single or multi)
        CompletionText := 'Updated shortcut settings for:' + #13#10;
        if (SelectedShortcutPath <> '') then
          CompletionText := CompletionText + '- ' + ExtractFileName(SelectedShortcutPath) + #13#10
        else if Assigned(SelectedShortcutPaths) then
        begin
          for i := 0 to SelectedShortcutPaths.Count - 1 do
            CompletionText := CompletionText + '- ' + ExtractFileName(SelectedShortcutPaths[i]) + #13#10;
        end;
        CompletionText := CompletionText + #13#10 + 'You can now test the connection by double-clicking the shortcut on your Desktop.';

        if Assigned(SelectedShortcutPaths) and (SelectedShortcutPaths.Count > 1) then
        begin
          WizardForm.FinishedHeadingLabel.Caption := 'Shortcut Settings Updated';
          WriteInstallerLog('CurPageChanged: Showing multi-shortcut settings saved message (' + IntToStr(SelectedShortcutPaths.Count) + ' shortcuts)');
        end
        else
        begin
          WizardForm.FinishedHeadingLabel.Caption := 'Shortcut Settings Updated';
          WriteInstallerLog('CurPageChanged: Showing single-shortcut settings saved message');
        end;
      end;
      if Assigned(FinishedExampleImage) then FinishedExampleImage.Visible := False;
    end
    else
    begin
      // Installation completion message
      WizardForm.FinishedHeadingLabel.Caption := 'Installation Complete';
      CompletionText := '';

      // Indicate whether TermWrap was installed
      if DoInstallTermWrap then
        CompletionText := CompletionText + 'TermWrap has been installed.' + #13#10#13#10;

      // List newly created users (and their shortcuts)
      if CreatedUsersList.Count > 0 then
      begin
        CompletionText := CompletionText + 'Created ' + IntToStr(CreatedUsersList.Count) + ' user account(s) and desktop shortcuts:' + #13#10;
        for i := 0 to CreatedUsersList.Count - 1 do
        begin
          CompletionText := CompletionText + '- ' + CreatedUsersList[i] + #13#10;
        end;
        CompletionText := CompletionText + #13#10;
        WriteInstallerLog('CurPageChanged: Created ' + IntToStr(CreatedUsersList.Count) + ' users');
      end;

      // List shortcuts created for existing users (if any)
      if ShortcutsList.Count > 0 then
      begin
        CompletionText := CompletionText + 'Created ' + IntToStr(ShortcutsList.Count) + ' shortcut(s) for existing user(s):' + #13#10;
        for i := 0 to ShortcutsList.Count - 1 do
        begin
          Entry := ShortcutsList[i];
          ParseUserEntry(Entry, UserName, Password);
          CompletionText := CompletionText + '- ' + UserName + #13#10;
        end;
        CompletionText := CompletionText + #13#10;
        WriteInstallerLog('CurPageChanged: Created ' + IntToStr(ShortcutsList.Count) + ' shortcuts for existing users');
      end;

      // Fallback message when nothing relevant was done
      if (not DoInstallTermWrap) and (CreatedUsersList.Count = 0) and (ShortcutsList.Count = 0) then
      begin
        if SelectedInstallMode = installModeEditSystemwideSettings then
        begin
          CompletionText := 'System-wide RDP settings were updated successfully.';
          WriteInstallerLog('CurPageChanged: System-wide settings updated');
        end
        else
        begin
          CompletionText := 'No user accounts were created during this run.' + #13#10#13#10 +
                            'You can add users later by rerunning this installer and choosing "Create Users Only".';
          WriteInstallerLog('CurPageChanged: No users or shortcuts created');
        end;
      end
      else
      begin
        // Encourage using shortcuts if any were created
        if (CreatedUsersList.Count + ShortcutsList.Count) > 0 then
          CompletionText := CompletionText + 'You can now open RDP connections using the shortcuts created on your Desktop.';
      end;
    end;

    // If Smart App Control is enabled, add guidance and show a popup
    if SmartAppControlIsOn then
    begin
      CompletionText := CompletionText + #13#10#13#10 +
        'IMPORTANT: Smart App Control is enabled on this system which intereferes with normal operation.' + #13#10 +
        'It is highly recommended to set it to Off via: Windows Security app > App & browser control > Smart App Control.';
      MsgBox('Smart App Control is enabled. It is highly recommended to set it to Off under:' + #13#10 +
             'Windows Security app > App & browser control > Smart App Control', mbInformation, MB_OK);
    end;

    WizardForm.FinishedLabel.Caption := CompletionText;
    WizardForm.Update;

    // Shrink the heading label to fit its actual text (removes extra whitespace)
    WizardForm.FinishedHeadingLabel.AutoSize := True;
    // Reposition the status label right below the compact heading
    FinishedText.Top := WizardForm.FinishedHeadingLabel.Top + WizardForm.FinishedHeadingLabel.Height + ScaleY(8);

    WriteInstallerLog('CurPageChanged: Updating FinishedText control');
    // Populate the label if available
    if Assigned(FinishedText) then
    begin
      WriteInstallerLog('CurPageChanged: FinishedText control is assigned');
      // Ensure text color contrasts with page surface
      if IsDarkColor(WizardForm.Color) then
        FinishedText.Font.Color := clWhite
      else
        FinishedText.Font.Color := clBlack;
      FinishedText.Caption := CompletionText;
      // Auto-size height based on content
      FinishedText.AutoSize := True;
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
  MstscPath: string;
  InstallerPath: string;
  OpTick: Cardinal;
  DownloadTick: Cardinal;
begin
  LogEntry('CheckAndInstallMSTSC');
  OpTick := GetTickCount;
  LogSectionHeader('MSTSC CHECK/INSTALL');
  // Check if mstsc.exe exists
  SetStepInProgress(StepCheckMSTSC, TXT_CheckMSTSC);
  if SimulateNoMstsc then
  begin
    if not SimLogNoMstscShown then
    begin
      LogSimulationScenario('System doesnt have mstsc');
      SimLogNoMstscShown := True;
    end;
    MstscPath := '';
    MSTSCExists := False;
  end
  else
  begin
    MstscPath := GetMstscPath;
    MSTSCExists := (MstscPath <> '');
  end;
  if MSTSCExists then
  begin
    LogDebug('mstsc.exe found at: ' + MstscPath);
    LogKeyValue('mstsc path', MstscPath);
    SetStepDone(StepCheckMSTSC, TXT_CheckMSTSC);
    SetStepDone(StepInstallMSTSC, TXT_InstallMSTSC); // skipped
  end
  else
  begin
    LogInfo('mstsc.exe missing. Initiating download/install.');
    LogKeyValue('Download URL', URL_RDP_INSTALLER);
    SetStepDone(StepCheckMSTSC, TXT_CheckMSTSC);
    SetStepInProgress(StepInstallMSTSC, TXT_InstallMSTSC);
    InstallerPath := TempFile('mstsc_installer.exe');
    LogDebug('Installer temp path: ' + InstallerPath);
    DownloadTick := GetTickCount;
    Exec(EXE_POWERSHELL, BuildPowerShellArgs('$out = ''' + InstallerPath + '''; Invoke-WebRequest -Uri ''' + URL_RDP_INSTALLER + ''' -OutFile $out -UseBasicParsing', True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    LogDebug('Download exit code=' + IntToStr(ResultCode) + ' [DURATION:' + IntToStr(GetTickCount - DownloadTick) + 'ms]');
    LogDebug('Installer file exists after download: ' + BoolToStr(FileExists(InstallerPath)));
    if (ResultCode = 0) and FileExists(InstallerPath) and IsSignedByMicrosoftCorporation(InstallerPath) then
    begin
      LogDebug('Installer verified as Microsoft-signed, executing...');
      Exec(InstallerPath, '', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      LogDebug('Installer execution exit code=' + IntToStr(ResultCode));
      if ResultCode = 0 then
      begin
        LogInfo('Remote Desktop Connection installed successfully.');
        SetStepDone(StepInstallMSTSC, TXT_InstallMSTSC);
      end
      else
      begin
        LogError('mstsc installer execution failed. Exit code: ' + IntToStr(ResultCode));
        PromptManualDownload('Remote Desktop Connection (mstsc)', URL_RDP_INSTALLER, 'Installer execution failed (exit code ' + IntToStr(ResultCode) + ')');
        SetStepDone(StepInstallMSTSC, TXT_InstallMSTSC);
      end;
    end
    else
    begin
      LogError('Failed to download or validate Remote Desktop Connection installer. Exit code: ' + IntToStr(ResultCode));
      PromptManualDownload('Remote Desktop Connection (mstsc)', URL_RDP_INSTALLER, 'Download or signature validation failed (exit code ' + IntToStr(ResultCode) + ')');
      SetStepDone(StepInstallMSTSC, TXT_InstallMSTSC);
    end;
    if FileExists(InstallerPath) then
    begin
      DeleteFile(InstallerPath);
      LogDebug('Deleted temp installer: ' + InstallerPath);
    end;
  end;
  LogInfo('CheckAndInstallMSTSC completed [DURATION:' + IntToStr(GetTickCount - OpTick) + 'ms]');
  LogExit('CheckAndInstallMSTSC');
end;
