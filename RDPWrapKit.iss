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

#define APP_VERSION_STRING "0.5.4"
#define APP_VERSION_FILEINFO "0.5.4.0"

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
WizardStyle=modern dynamic
SetupIconFile="assets\RDPWrapKitIcon.ico"

[Files]
; Icon file always extracted to temp for welcome page display.
; TermWrap files only copied when DoInstallTermWrap = True (checked via ShouldInstallFiles).
Source: "third_party\termwrap_release\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs; Check: ShouldInstallFiles
Source: "assets\RDPWrapKitIcon.bmp"; DestDir: "{tmp}"; Flags: ignoreversion dontcopy
Source: "assets\rdp_edit_save.bmp"; DestDir: "{tmp}"; Flags: ignoreversion skipifsourcedoesntexist dontcopy



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
procedure OnShortcutRadioClick(Sender: TObject); forward;
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
function BoolToStr(Value: Boolean): string; forward;
procedure EnsureTermServiceRunsAsNetworkService; forward;
procedure EnsureUmRdpServiceAutomatic; forward;
procedure InitInstallerLog; forward;
procedure WriteInstallerLog(const Msg: string); forward;
procedure LogSectionHeader(const Title: string); forward;
procedure LogKeyValue(const KeyName, KeyValue: string); forward;
var
  Page_InstallOptions: TWizardPage;
  WelcomePage: TWizardPage;
  UserPage: TInputQueryWizardPage;
  // AdvancedPage removed - single create-shortcuts page retained
  EditSystemwideSettingsPage: TWizardPage;  // Main Edit System-wide settings page
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
  rbEditShortcutSettings: TRadioButton;
  CreateRdpShortcutsGroup: TPanel;
  EditShortcutPage: TWizardPage;
  Page_ShortcutSettings: TWizardPage;
  DesktopRdpFiles: TStringList;
  ShortcutRadioButtons: array of TRadioButton;
  CurrentShortcutPage: Integer;
  ShortcutsPerPage: Integer;
  ShortcutPrevButton: TButton;
  ShortcutNextButton: TButton;
  ShortcutPageLabel: TLabel;
  ShortcutHeaderLabel: TLabel;
  ShortcutEmptyLabel: TLabel;
  EditShortcutControlsBuilt: Boolean;
  SelectedShortcutIndex: Integer;
  SelectedShortcutPath: string;
  FinishedExampleImage: TBitmapImage;
  
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
  lblShortcutTips: TLabel;
  lblShortcutEditingFile: TLabel;
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
  TXT_ConfigureService = 'Configure TermWrap service';
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
  BUILD_FINGERPRINT = 'stabledebug-v21-ps-dquote-fix';

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
    Msg := 'A newer version is available: ' + LatestVersion + #13#10
         + 'You are about to install: ' + CurrentVersion + #13#10#13#10
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
  PipePos := Pos('|', Entry);
  UserName := Copy(Entry, 1, PipePos - 1);
  Password := Copy(Entry, PipePos + 1, Length(Entry));
end;

// Return a version of a pipe-delimited user entry with the password obscured
function MaskPasswordInEntry(const Entry: string): string;
var
  PipePos: Integer;
begin
  PipePos := Pos('|', Entry);
  if PipePos > 0 then
    Result := Copy(Entry, 1, PipePos) + '*****'
  else
    Result := Entry;
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
  BaseDir := ExpandConstant('{localappdata}\RDPWrapKit');
  if (not DirExists(BaseDir)) and (not CreateDir(BaseDir)) then
  begin
    WriteInstallerLog('WARNING: Could not create debug work base directory: ' + BaseDir);
    Result := ExpandConstant('{tmp}');
    exit;
  end;

  Result := BaseDir + '\DebugLogs';
  if (not DirExists(Result)) and (not CreateDir(Result)) then
  begin
    WriteInstallerLog('WARNING: Could not create debug work directory: ' + Result);
    Result := ExpandConstant('{tmp}');
    exit;
  end;

  WriteInstallerLog('Debug work directory ready: ' + Result);
end;

function DebugLogFile(const FileName: string): string;
begin
  Result := EnsureDebugWorkDir + '\' + FileName;
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
  // Double-quote the value so Windows CreateProcess (CommandLineToArgvW) strips the
  // surrounding quotes correctly when PowerShell runs in -File mode.
  // Single quotes have no special meaning in CreateProcess and would be passed as literals.
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
begin
  ScriptPath := TempFile(ScriptBaseName);
  SaveStringToFile(ScriptPath, ScriptContent, False);
  WriteInstallerLog('PowerShell File: ' + ScriptPath + ' ' + MaskPasswordsInString(ExtraParams));
  Result := Exec(EXE_POWERSHELL, BuildPowerShellFileArgs(ScriptPath, ExtraParams, Hidden), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  DeleteFile(ScriptPath);
end;

function ExecSavedPowerShellDebugScriptParams(const ScriptTag, UserName, ScriptContent, ExtraParams: string; Hidden: Boolean; var ResultCode: Integer): Boolean;
var
  ScriptPath: string;
begin
  ScriptPath := TempFile(ScriptTag + '_' + SanitizeFileName(UserName) + '.ps1');
  SaveStringToFile(ScriptPath, ScriptContent, False);
  WriteInstallerLog('PowerShell File: ' + ScriptPath + ' ' + MaskPasswordsInString(ExtraParams));
  Result := Exec(EXE_POWERSHELL, BuildPowerShellFileArgs(ScriptPath, ExtraParams, Hidden), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  if not Result then
    WriteInstallerLog('ERROR: Failed to launch PowerShell script: code=' + IntToStr(ResultCode) + ' message=' + SysErrorMessage(ResultCode));
  DeleteFile(ScriptPath);
end;

function BuildAddGroupMemberPowerShellScript(const GroupName, UserName, OutPath, SuccessTag: string): string;
begin
  Result :=
    'param([string]$GroupName, [string]$UserName, [string]$OutPath, [string]$SuccessTag)' + #13#10 +
    '$ErrorActionPreference = ''Stop''' + #13#10 +
    'try {' + #13#10 +
    '  Add-LocalGroupMember -Group $GroupName -Member $UserName -ErrorAction Stop' + #13#10 +
    '  @($SuccessTag, (''Group={0}'' -f $GroupName), (''User={0}'' -f $UserName)) | Out-File -FilePath $OutPath -Encoding UTF8' + #13#10 +
    '  exit 0' + #13#10 +
    '} catch {' + #13#10 +
    '  @(' + #13#10 +
    '    ''ADD_GROUP_FAIL'',' + #13#10 +
    '    (''Group={0}'' -f $GroupName),' + #13#10 +
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
  Result := False;
  OutPath := TempFile('grp_members_' + SanitizeFileName(GroupName) + '.txt');

  // net localgroup lists members one per line after a "---" separator line.
  // Members may appear as DOMAIN\username or bare username.
  Exec('cmd.exe', '/c net localgroup ' + QuoteExeArg(GroupName) + ' > ' + QuoteExeArg(OutPath) + ' 2>&1',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  if not FileExists(OutPath) then
    exit;

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
      // Strip domain prefix if present (DOMAIN\username -> username)
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
end;

// Verify file is authenticode-signed by Microsoft Corporation
function IsSignedByMicrosoftCorporation(const FilePath: string): Boolean;
var
  OutText: string;
begin
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
  WriteInstallerLog('Signature details: ' + OutText);
  Result := Pos('|RESULT=OK', UpperCase(Trim(OutText))) > 0;
  if Result then
    WriteInstallerLog('Signature verdict: Microsoft publisher validation passed')
  else
    WriteInstallerLog('Signature verdict: Microsoft publisher validation failed');
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
begin
  PSArgs := BuildPowerShellArgs(Command, True);
  // Log command and run (mask any embedded passwords)
  WriteInstallerLog('PowerShell Hidden: ' + MaskPasswordsInString(PSArgs));
  Result := Exec(EXE_POWERSHELL, PSArgs, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  WriteInstallerLog('PowerShell exitcode=' + IntToStr(ResultCode));
end;

// Execute a PowerShell command and capture stdout to a temp file, returning
// the trimmed output as a string. Returns empty string on failure.
function GetPSOutput(const Command: string): string;
var
  PSPath: string;
  RC: Integer;
  SL: TStringList;
begin
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
    finally
      SL.Free;
      DeleteFile(PSPath);
    end;
  end;
end;

procedure SignRdpFile(const RdpPath: string);
var
  PSCommand: string;
  ResultCode: Integer;
begin
  // Ensure cert exists and sign the .rdp file using rdpsign.exe
  PSCommand :=
    '$subjectName = ''CN=RDPWrapKit: Only trust if connecting to 127.0.0.2''; ' +
    '$existing = Get-ChildItem "Cert:\\LocalMachine\\My" | Where-Object { $_.Subject -eq $subjectName } | Select-Object -First 1; ' +
    'if ($existing) { $thumb = $existing.Thumbprint } else { ' +
      '$c = New-SelfSignedCertificate -Subject $subjectName -CertStoreLocation "Cert:\\LocalMachine\\My" -KeyUsage DigitalSignature -Type CodeSigningCert -NotAfter (Get-Date).AddYears(10); ' +
      '$thumb = $c.Thumbprint; $tmp = Join-Path $env:TEMP "rdpwrapkit.cer"; Export-Certificate -Cert ("Cert:\\LocalMachine\\My\\" + $thumb) -FilePath $tmp; Import-Certificate -FilePath $tmp -CertStoreLocation "Cert:\\LocalMachine\\Root"; Remove-Item $tmp -Force } ; ' +
    'try { & rdpsign.exe /sha256 $thumb "' + RdpPath + '"; exit $LASTEXITCODE } catch { exit 1 }';

  Exec(EXE_POWERSHELL, BuildPowerShellArgs(PSCommand, True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  if ResultCode = 0 then
    WriteInstallerLog('SignRdpFile: signed ' + RdpPath)
  else
    WriteInstallerLog('SignRdpFile: failed to sign ' + RdpPath + ' (exit=' + IntToStr(ResultCode) + ')');
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

  WriteInstallerLog('SystemInfo: Build.UBR=' + BuildNumber + '.' + IntToStr(UBR));
  if DisplayVersion <> '' then
    WriteInstallerLog('SystemInfo: Version=' + DisplayVersion);
  if BuildNumber <> '' then
    WriteInstallerLog('SystemInfo: Build=' + BuildNumber + ' (UBR ' + IntToStr(UBR) + ')');
  if EditionID <> '' then
    WriteInstallerLog('SystemInfo: Edition=' + EditionID);
  if InstallLang <> '' then
    WriteInstallerLog('SystemInfo: InstallLanguage=' + InstallLang);
  WriteInstallerLog('SystemInfo: Arch=' + Arch);
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

function RepeatChar(const Ch: string; const Count: Integer): string;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Count do
    Result := Result + Ch;
end;

procedure LogSectionHeader(const Title: string);
begin
  WriteInstallerLog('+' + RepeatChar('-', 78) + '+');
  WriteInstallerLog('| ' + Title);
  WriteInstallerLog('+' + RepeatChar('-', 78) + '+');
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
  else if Assigned(EditSystemwideSettingsPage) and (PageID = EditSystemwideSettingsPage.ID) then Result := 'Custom: Edit System-wide RDP Settings'
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
  WriteInstallerLog('  - ' + KeyName + ': ' + KeyValue);
end;

procedure DumpTextFileToLog(const HeaderText, FilePath: string);
var
  Tmp: TStringList;
  k: Integer;
begin
  if not FileExists(FilePath) then
  begin
    WriteInstallerLog('WARNING: ' + HeaderText + ' file not found: ' + FilePath);
    exit;
  end;

  Tmp := TStringList.Create;
  try
    try
      Tmp.LoadFromFile(FilePath);
      WriteInstallerLog(HeaderText + ' (' + IntToStr(Tmp.Count) + ' lines):');
      for k := 0 to Tmp.Count - 1 do
        WriteInstallerLog('DEBUG: ' + Tmp[k]);
    except
      WriteInstallerLog('WARNING: Failed to read debug output file: ' + FilePath);
    end;
  finally
    Tmp.Free;
  end;
end;

procedure LogPasswordPipeline(const StageName, UserName, Password: string);
begin
  if PASSWORD_PIPELINE_DIAG = 0 then
    exit;
  WriteInstallerLog('PASSWORD_DIAG [' + StageName + '] user=' + UserName + ' :: details=<redacted>');
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
    WriteInstallerLog('PASSWORD_DIAG [' + StageName + '] encrypted file missing: ' + FilePath);
    exit;
  end;

  if not LoadStringFromFile(FilePath, EncRaw) then
  begin
    WriteInstallerLog('PASSWORD_DIAG [' + StageName + '] failed reading encrypted file: ' + FilePath);
    exit;
  end;

  EncText := Trim(String(EncRaw));
  WriteInstallerLog('PASSWORD_DIAG [' + StageName + '] encLen=' + IntToStr(Length(EncText)) +
    ' | encPrefix=' + Copy(EncText, 1, 24));
end;

procedure PreserveDebugLogFileToDesktop(const FilePath: string);
var
  DestDir: string;
  DestPath: string;
  BaseName: string;
begin
  if PRESERVE_USER_CREATE_DEBUG_LOGS = 0 then
    exit;

  if not FileExists(FilePath) then
    exit;

  DestDir := ExpandConstant('{userdesktop}\RDPWrapKit_DebugLogs');
  if (not DirExists(DestDir)) and (not CreateDir(DestDir)) then
  begin
    WriteInstallerLog('WARNING: Could not create debug log folder: ' + DestDir);
    exit;
  end;

  BaseName := ChangeFileExt(ExtractFileName(FilePath), '');
  DestPath := DestDir + '\' + BaseName + '_' + IntToStr(GetTickCount) + '.log';
  if CopyFile(FilePath, DestPath, False) then
    WriteInstallerLog('Saved debug user-create log: ' + DestPath)
  else
    WriteInstallerLog('WARNING: Failed to save debug user-create log copy for ' + FilePath);
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
    WriteInstallerLog('Deferred cleanup registered for Finish: ' + FilePath);
  end;
end;

procedure CleanupPendingDebugFiles;
var
  i: Integer;
  P: string;
begin
  if not Assigned(PendingDebugCleanupFiles) then
    exit;

  if CLEANUP_DEBUG_FILES_ON_FINISH = 0 then
  begin
    LogSectionHeader('FINISH CLEANUP: DEFERRED DEBUG FILES');
    WriteInstallerLog('Deferred debug cleanup skipped by configuration; files retained for troubleshooting');
    LogKeyValue('Queued files retained', IntToStr(PendingDebugCleanupFiles.Count));
    PendingDebugCleanupFiles.Clear;
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
        WriteInstallerLog('Deleted deferred debug file: ' + P)
      else
        WriteInstallerLog('WARNING: Failed to delete deferred debug file: ' + P);
    end
    else
      WriteInstallerLog('Deferred debug file already missing: ' + P);
  end;

  PendingDebugCleanupFiles.Clear;
end;

procedure PromptManualDownload(const ComponentName, Url, Reason: string);
var
  Choice: Integer;
  RC: Integer;
begin
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
    WriteInstallerLog('Manual download prompt: user chose YES for ' + ComponentName);
    if not ShellExec('', Url, '', '', SW_SHOWNORMAL, ewNoWait, RC) then
      WriteInstallerLog('Manual download launch failed for ' + ComponentName + ', ShellExec rc=' + IntToStr(RC))
    else
      WriteInstallerLog('Manual download launch succeeded for ' + ComponentName);
  end
  else
  begin
    WriteInstallerLog('Manual download prompt: user chose NO for ' + ComponentName);
  end;
end;

procedure InitInstallerLog;
begin
  InstallLogPath := ExpandConstant(INSTALL_LOG_PATH);
  try
    SaveStringToFile(InstallLogPath, GetTimestampString + ' +' + RepeatChar('=', 78) + '+' + #13#10, False);
    SaveStringToFile(InstallLogPath, GetTimestampString + ' | RDPWrapKit Installer Log' + #13#10, True);
    SaveStringToFile(InstallLogPath, GetTimestampString + ' | Session started (UTC)' + #13#10, True);
    SaveStringToFile(InstallLogPath, GetTimestampString + ' +' + RepeatChar('=', 78) + '+' + #13#10, True);
  except
  end;
  LogSectionHeader('ENVIRONMENT SNAPSHOT');
  LogSystemInfo;
end;

procedure WriteInstallerLog(const Msg: string);
begin
  try
    SaveStringToFile(InstallLogPath, GetTimestampString + ' ' + BeautifyLogMessage(Msg) + #13#10, True);
  except
  end;
end;

// Run any process hidden and return its exit code
function RunHidden(const FileName, Params: string): Integer;
var
  RC: Integer;
begin
  RC := 0;
  WriteInstallerLog('Exec: ' + FileName + ' ' + MaskCommandForLog(FileName, Params));
  Exec(FileName, Params, '', SW_HIDE, ewWaitUntilTerminated, RC);
  WriteInstallerLog('ExitCode: ' + IntToStr(RC) + ' for ' + FileName);
  Result := RC;
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
  WriteInstallerLog('SIMULATED SCENARIO: ' + ScenarioText);
  Log('SIMULATED SCENARIO: ' + ScenarioText);
end;

function RunNetHidden(const Params: string): Integer;
begin
  if SimulateNetFailPowerShell then
  begin
    if SimulateNetFailPowerShell and (not SimLogNetPsShown) then
    begin
      LogSimulationScenario('System fails on net.exe commands and uses PowerShell fallback');
      SimLogNetPsShown := True;
    end;
    WriteInstallerLog('Simulation: forcing net.exe failure for params: ' + MaskCommandForLog('net.exe', Params));
    Result := 1;
    exit;
  end;

  Result := RunHidden('net.exe', Params);
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
  WriteInstallerLog('RunPSHiddenCode: ' + MaskPasswordsInString(Command));
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
begin
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
begin
  // Ensure Defender exclusions are scoped to the two runtime DLLs only
  ExecPowerShellHidden(
    '$paths = @(''' + ExpandConstant('{app}\TermWrap.dll') + ''',''' + ExpandConstant('{app}\Zydis.dll') + '''); ' +
    'try { $p = Get-MpPreference; foreach ($path in $paths) { if (-not ($p.ExclusionPath -contains $path)) { Add-MpPreference -ExclusionPath $path } } } catch { }',
    ResultCode);
end;

procedure RemoveDefenderExclusionForApp;
var
  ResultCode: Integer;
begin
  // Remove Defender exclusions for the two runtime DLLs during uninstall
  ExecPowerShellHidden(
    '$paths = @(''' + ExpandConstant('{app}\TermWrap.dll') + ''',''' + ExpandConstant('{app}\Zydis.dll') + '''); ' +
    'try { foreach ($path in $paths) { Remove-MpPreference -ExclusionPath $path } } catch { }',
    ResultCode);
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

// Shared helper: query a string value, fix it via sc.exe if wrong, then verify.
procedure EnsureServiceStringConfig(const RegKey, ValueName, ScArgs, TargetValue, CheckDesc, FixDesc: string);
var
  Current: string;
  RC: Integer;
begin
  Current := '';
  if RegQueryStringValue(HKLM, RegKey, ValueName, Current) then
    WriteInstallerLog(CheckDesc + ': current ' + ValueName + '=' + Current)
  else
    WriteInstallerLog(CheckDesc + ': current ' + ValueName + '=<missing>');

  if CompareText(Trim(Current), TargetValue) = 0 then
    Exit;

  RC := RunHidden('sc.exe', ScArgs);
  if RC = 0 then
    WriteInstallerLog(FixDesc + ': success')
  else
    WriteInstallerLog('WARNING: ' + FixDesc + ' failed: sc.exe exit=' + IntToStr(RC));

  Sleep(SLEEP_SHORT);
  Current := '';
  if RegQueryStringValue(HKLM, RegKey, ValueName, Current) then
    WriteInstallerLog(CheckDesc + ' (post-fix): ' + ValueName + '=' + Current)
  else
    WriteInstallerLog('WARNING: ' + CheckDesc + ' (post-fix): ' + ValueName + ' could not be read');
end;

// Shared helper: query a DWord value, fix it via sc.exe if wrong, then verify.
procedure EnsureServiceDWordConfig(const RegKey, ValueName, ScArgs: string; TargetValue: Cardinal; const CheckDesc, FixDesc: string);
var
  Current: Cardinal;
  RC: Integer;
begin
  Current := $FFFFFFFF;
  if RegQueryDWordValue(HKLM, RegKey, ValueName, Current) then
    WriteInstallerLog(CheckDesc + ': current ' + ValueName + '=' + IntToStr(Current))
  else
    WriteInstallerLog(CheckDesc + ': current ' + ValueName + '=<unable to read>');

  if Current = TargetValue then
    Exit;

  RC := RunHidden('sc.exe', ScArgs);
  if RC = 0 then
    WriteInstallerLog(FixDesc + ': success')
  else
    WriteInstallerLog('WARNING: ' + FixDesc + ' failed: sc.exe exit=' + IntToStr(RC));

  Sleep(SLEEP_SHORT);
  if RegQueryDWordValue(HKLM, RegKey, ValueName, Current) then
    WriteInstallerLog(CheckDesc + ' (post-fix): ' + ValueName + '=' + IntToStr(Current))
  else
    WriteInstallerLog('WARNING: ' + CheckDesc + ' (post-fix): ' + ValueName + ' could not be read');
end;

procedure EnsureTermServiceRunsAsNetworkService;
begin
  EnsureServiceStringConfig(
    REG_TERMSERVICE, 'ObjectName',
    'config TermService obj= "NT AUTHORITY\NetworkService" password= ""',
    'NT AUTHORITY\NetworkService',
    'TermService account check',
    'TermService account fix: set ObjectName to NT AUTHORITY\NetworkService');
end;

// Wrapper that calls StartTermServiceEx and discards the exit code
procedure StartTermService;
begin
  StartTermServiceEx;
end;

procedure EnsureUmRdpServiceAutomatic;
begin
  // Start type 2 = Automatic
  EnsureServiceDWordConfig(
    REG_UMRDPSERVICE, 'Start',
    'config UmRdpService start=auto',
    2,
    'UmRdpService startup type check',
    'UmRdpService startup type fix: set Start=2 (Automatic)');
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
begin
  PSPath := ExpandConstant(TEMP_LOCAL_USERS);

  // Get-LocalUser is available on all Win11 Home/Pro editions. It is the only
  // reliable way to filter ghost/orphaned accounts via PrincipalSource.
  // For Microsoft-linked accounts, the connected email is read from the
  // IdentityStore registry cache. Inlined via -Command; single-quotes used
  // throughout to avoid conflicts with the outer double-quote wrapper.
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

  if not FileExists(PSPath) then
    exit;

  Parts := TStringList.Create;
  try
    Parts.LoadFromFile(PSPath);
    for i := 0 to Parts.Count - 1 do
    begin
      Line := Trim(Parts[i]);
      if Line = '' then
        continue;
      // Split "username|displayname"
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
        continue;
      UsersList.Add(UserName);
      DisplayList.Add(DisplayName);
    end;
  finally
    Parts.Free;
    DeleteFile(PSPath);
  end;
end;

function GetDesktopRdpFiles: TStringList;
var
  FilesList: TStringList;
  FindRec: TFindRec;
  DesktopPattern: string;
begin
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

procedure OnShortcutRadioClick(Sender: TObject);
begin
  SelectedShortcutIndex := TRadioButton(Sender).Tag;
end;

procedure BuildShortcutEditorControls;
var
  i: Integer;
begin
  if EditShortcutControlsBuilt then
    exit;

  if ShortcutsPerPage = 0 then
    ShortcutsPerPage := 8;

  ShortcutHeaderLabel := TLabel.Create(EditShortcutPage);
  ShortcutHeaderLabel.Parent := EditShortcutPage.Surface;
  ShortcutHeaderLabel.Left := ScaleX(20);
  ShortcutHeaderLabel.Top := ScaleY(12);
  ShortcutHeaderLabel.Caption := 'Desktop .rdp shortcuts';
  ShortcutHeaderLabel.Font.Style := [fsBold];

  ShortcutEmptyLabel := TLabel.Create(EditShortcutPage);
  ShortcutEmptyLabel.Parent := EditShortcutPage.Surface;
  ShortcutEmptyLabel.Left := ScaleX(20);
  ShortcutEmptyLabel.Top := ShortcutHeaderLabel.Top + ScaleY(28);
  ShortcutEmptyLabel.Caption := 'No .rdp files were found on your Desktop.';
  ShortcutEmptyLabel.Visible := False;

  SetLength(ShortcutRadioButtons, DesktopRdpFiles.Count);
  for i := 0 to DesktopRdpFiles.Count - 1 do
  begin
    ShortcutRadioButtons[i] := TRadioButton.Create(EditShortcutPage);
    ShortcutRadioButtons[i].Parent := EditShortcutPage.Surface;
    ShortcutRadioButtons[i].Left := ScaleX(20);
    ShortcutRadioButtons[i].Width := EditShortcutPage.SurfaceWidth - ScaleX(40);
    ShortcutRadioButtons[i].Caption := ExtractFileName(DesktopRdpFiles[i]);
    ShortcutRadioButtons[i].Tag := i;
    ShortcutRadioButtons[i].OnClick := @OnShortcutRadioClick;
    ShortcutRadioButtons[i].Visible := False;
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

procedure UpdateShortcutPageDisplay;
var
  i, PageCount, StartIdx, EndIdx, VisIndex, BaseTop, RowHeight: Integer;
begin
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
    ShortcutRadioButtons[i].Visible := (i >= StartIdx) and (i <= EndIdx);
    if ShortcutRadioButtons[i].Visible then
    begin
      ShortcutRadioButtons[i].Top := BaseTop + VisIndex * RowHeight;
      ShortcutRadioButtons[i].Checked := (i = SelectedShortcutIndex);
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
  if CurrentShortcutPage > 0 then
    Dec(CurrentShortcutPage);
  UpdateShortcutPageDisplay;
end;

procedure OnNextShortcutPageClick(Sender: TObject);
var
  PageCount: Integer;
begin
  PageCount := (DesktopRdpFiles.Count + ShortcutsPerPage - 1) div ShortcutsPerPage;
  if CurrentShortcutPage < PageCount - 1 then
    Inc(CurrentShortcutPage);
  UpdateShortcutPageDisplay;
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

procedure OnCreateRdpShortcutsClick(Sender: TObject);
begin
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
end;

  procedure OnInstallModeChange(Sender: TObject);
begin
  // Always show controls, but only enable them for Install mode
  if Assigned(chkInstallTermWrap) then chkInstallTermWrap.Enabled := Assigned(rbInstall) and rbInstall.Checked;
  if Assigned(chkCreateRdpShortcuts) then chkCreateRdpShortcuts.Enabled := Assigned(rbInstall) and rbInstall.Checked;
  if Assigned(CreateRdpShortcutsGroup) then CreateRdpShortcutsGroup.Enabled := Assigned(rbInstall) and rbInstall.Checked;
  if Assigned(rbCreateUsers) then rbCreateUsers.Enabled := Assigned(rbInstall) and rbInstall.Checked and chkCreateRdpShortcuts.Checked;
  if Assigned(rbUseExistingUsers) then rbUseExistingUsers.Enabled := Assigned(rbInstall) and rbInstall.Checked and chkCreateRdpShortcuts.Checked;
end;

procedure OnUseAllMonitorsClick(Sender: TObject);
begin
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
end;

procedure OnFullScreenClick(Sender: TObject);
begin
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
end;

procedure OnResolutionChange(Sender: TObject);
var
  IsCustom: Boolean;
begin
  if not Assigned(cboResolution) then exit;
  IsCustom := (cboResolution.ItemIndex >= 0) and
              (cboResolution.Items[cboResolution.ItemIndex] = 'Custom');
  if Assigned(lblCustomWidth)  then lblCustomWidth.Visible  := IsCustom;
  if Assigned(edtCustomWidth)  then edtCustomWidth.Visible  := IsCustom;
  if Assigned(lblCustomHeight) then lblCustomHeight.Visible := IsCustom;
  if Assigned(edtCustomHeight) then edtCustomHeight.Visible := IsCustom;
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


procedure OnViewLogButtonClick(Sender: TObject);
var
  DestName: string;
  Saved: Boolean;
begin
  WriteInstallerLog('User clicked Save Install Log button');
  DestName := ExpandConstant('{userdesktop}\RDPWrapKit_install.log');
  Saved := CopyFile(InstallLogPath, DestName, False);
  if Saved then
    MsgBox('Install log saved to:' + #13#10 + DestName, mbInformation, MB_OK)
  else
    MsgBox('Failed to save install log to the Desktop location.', mbError, MB_OK);
end;

procedure OnPasswordResetLinkClick(Sender: TObject);
var
  ResultCode: Integer;
begin
  Exec('control.exe', 'userpasswords2', '', SW_SHOW, ewNoWait, ResultCode);
end;

procedure BuildCreateShortcutsControls;
var
  i: Integer;
  TopPos: Integer;
  BottomPos: Integer;
begin
  // Avoid building controls multiple times (prevents duplicate buttons/labels)
  if CreateShortcutsControlsBuilt then
    exit;
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
  if LocalUsersList.Count = 0 then
    exit;

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
  if CurrentUserPage > 0 then
    Dec(CurrentUserPage);
  UpdateUsersPageDisplay;
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
var
  i: Integer;
  c: Char;
begin
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
  // Fast local credential check via LogonUser; avoids slow PowerShell/WinRM
  Token := 0;
  Result := LogonUser(UserName, '.', Password, 2, 0, Token);
  if Token <> 0 then
    CloseHandle(Token);
end;

function ShouldInstallFiles: Boolean;
begin
  // Only install bundled TermWrap files when Install TermWrap is selected or when the
  // user explicitly selected "Install TermWrap" on the welcome/options page.
  Result := DoInstallTermWrap;
end;

function ShouldApplyRegistryEntries: Boolean;
begin
  // Edit Shortcut Settings mode should only launch mstsc /edit and avoid
  // unrelated installer-side registry changes.
  Result := SelectedInstallMode <> installModeEditShortcuts;
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
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

  // Skip Ready page - no need to show install path
  if PageID = wpReady then
    Result := True;

  // Show Edit System-wide RDP settings page only in that mode
  if (PageID = EditSystemwideSettingsPage.ID) and (SelectedInstallMode <> installModeEditSystemwideSettings) then
    Result := True;

  // Show Page_ShowRDPInfo only in that mode
  if Assigned(Page_ShowRDPInfo) and (PageID = Page_ShowRDPInfo.ID) and (SelectedInstallMode <> installModeShowRDPInfo) then
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
  // Securely delete temporary files that contained sensitive data
  DeleteFile(TempFile('enc_' + UserName + '.txt'));
  DeleteFile(TempFile('create_rdp_' + UserName + '.ps1'));
end;

function EncryptPasswordToFile(const Password, UserName: string): string;
var
  ResultCode: Integer;
  EncPath: string;
  Cmd: string;
begin
  EncPath := TempFile('enc_' + UserName + '.txt');

  // Use inline PowerShell command (no temporary .ps1 file) to avoid script-file stalls.
  Cmd :=
    '$pw = ''' + PSSingleQuote(Password) + ''' | ConvertTo-SecureString -AsPlainText -Force; ' +
    '$encPw = ConvertFrom-SecureString $pw; ' +
    '[System.IO.File]::WriteAllText(''' + PSSingleQuote(EncPath) + ''', $encPw)';

  ResultCode := -1;
  Exec(EXE_POWERSHELL, BuildPowerShellArgs(Cmd, True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  if (ResultCode = 0) and FileExists(EncPath) then
  begin
    Result := EncPath;
  end
  else
  begin
    if FileExists(EncPath) then
      DeleteFile(EncPath);
    WriteInstallerLog('WARNING: EncryptPasswordToFile failed for user ' + UserName + ', proceeding without embedded password');
    Result := '';
  end;
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
procedure GetShortcutDisplaySettings(var ScreenModeId, DesktopWidth, DesktopHeight, UseMultiMon, AudioMode, RedirectClipboard: Integer);
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
end;

function WriteRDPFileDirect(const UserName, RDPPath, EncPath: string): Boolean;
var
  ScreenModeId: Integer;
  DesktopWidth: Integer;
  DesktopHeight: Integer;
  UseMultiMon: Integer;
  AudioMode: Integer;
  RedirectClipboard: Integer;
  DisableWallpaper, AllowFontSmooth, AllowComposition: Integer;
  DisableFullWindowDrag, DisableMenuAnims, DisableThemes: Integer;
  SL: TStringList;
  EncTextRaw: AnsiString;
  EncText: string;
  HasEnc: Boolean;
begin
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

  GetShortcutDisplaySettings(ScreenModeId, DesktopWidth, DesktopHeight, UseMultiMon, AudioMode, RedirectClipboard);
  GetExperienceSettings(DisableWallpaper, AllowFontSmooth, AllowComposition, DisableFullWindowDrag, DisableMenuAnims, DisableThemes);

  SL := TStringList.Create;
  try
    SL.Add('full address:s:' + RDP_LOOPBACK_IP);
    SL.Add('username:s:' + UserName);
    SL.Add('screen mode id:i:' + IntToStr(ScreenModeId));
    SL.Add('desktopwidth:i:' + IntToStr(DesktopWidth));
    SL.Add('desktopheight:i:' + IntToStr(DesktopHeight));
    SL.Add('use multimon:i:' + IntToStr(UseMultiMon));
    SL.Add('session bpp:i:32');
    SL.Add('smart sizing:i:1');
    SL.Add('dynamic resolution:i:1');
    SL.Add('autoreconnection enabled:i:1');
    SL.Add('compression:i:1');
    SL.Add('keyboardhook:i:1');
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
  DisableWallpaper, AllowFontSmooth, AllowComposition: Integer;
  DisableFullWindowDrag, DisableMenuAnims, DisableThemes: Integer;
begin
  GetShortcutDisplaySettings(ScreenModeId, DesktopWidth, DesktopHeight, UseMultiMon, AudioMode, RedirectClipboard);
  GetExperienceSettings(DisableWallpaper, AllowFontSmooth, AllowComposition, DisableFullWindowDrag, DisableMenuAnims, DisableThemes);

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
    '  $rdp += "full address:s:' + RDP_LOOPBACK_IP + '"' + #13#10 +
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
    '  $rdp += "keyboardhook:i:1"' + #13#10 +
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
begin
  RDPPath := ExpandConstant('{userdesktop}\' + UserName + '.rdp');
  ScriptPath := TempFile('create_rdp_' + UserName + '.ps1');

  if PASSWORD_PIPELINE_DIAG <> 0 then
  begin
    LogSectionHeader('PASSWORD PIPELINE DEBUG');
    LogKeyValue('User', UserName);
    LogKeyValue('Creation Source', CreationSource);
    LogPasswordPipeline('SHORTCUT_INPUT', UserName, Password);
  end;

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
  if EncPath <> '' then
    LogEncryptedFileSummary('AFTER_ENCRYPT', EncPath);

  // Direct write path avoids PowerShell hangs in shortcut generation.
  if WriteRDPFileDirect(UserName, RDPPath, EncPath) then
  begin
    // Sign the .rdp file immediately after direct write
    SignRdpFile(RDPPath);
    SecureCleanupTempFiles(UserName);
    exit;
  end
  else
  begin
    WriteInstallerLog('CreateRDPShortcut: Direct write failed, falling back to PowerShell script path');
  end;

  // Generate and execute PowerShell script
  WriteInstallerLog('CreateRDPShortcut: Generating RDP PowerShell script');
  PowerShellScript := GenerateRDPPowerShellScript(UserName, RDPPath, EncPath);
  SaveStringToFile(ScriptPath, PowerShellScript, False);
  WriteInstallerLog('CreateRDPShortcut: Script saved to ' + ScriptPath);
  
  // Wrap script execution in a timed job to avoid installer hangs
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
  WriteInstallerLog('CreateRDPShortcut: Executing PowerShell script with timeout wrapper');
  Exec(EXE_POWERSHELL,
    BuildPowerShellFileArgs(
      RunnerPath,
      BuildPSNamedParam('TargetScript', ScriptPath) + ' ' + BuildPSNamedParam('EncPath', EncPath) + ' -TimeoutSec 30',
      True),
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  WriteInstallerLog('CreateRDPShortcut: Wrapped PowerShell script exit code=' + IntToStr(ResultCode));
  DeleteFile(RunnerPath);
  
  if ResultCode <> 0 then
  begin
    Log('WARNING: RDP file creation failed with exit code = ' + IntToStr(ResultCode));
    WriteInstallerLog('CreateRDPShortcut: WARNING - RDP file creation failed with exit code=' + IntToStr(ResultCode));
  end
  else
  begin
    WriteInstallerLog('CreateRDPShortcut: RDP file created successfully');
    // Sign the created RDP file so saved shortcuts are trusted by our cert
    SignRdpFile(RDPPath);
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

    ValidationError := IsValidUsername(UserName);
    if ValidationError <> '' then
    begin
      WriteInstallerLog('ERROR: Skipping user due to invalid username input: ' + ValidationError + ' | User=' + UserName);
      continue;
    end;
    ValidationError := IsValidPassword(Password);
    if ValidationError <> '' then
    begin
      WriteInstallerLog('ERROR: Skipping user due to invalid password input: ' + ValidationError + ' | User=' + UserName);
      continue;
    end;

    WizardForm.StatusLabel.Caption := 'Creating user account (' + IntToStr(i + 1) + ' of ' + IntToStr(UsersList.Count) + '): ' + UserName;
    WriteInstallerLog('Creating user: ' + UserName);
    UserCreateOutputAlreadyLogged := False;
    UserCreatePath := 'NET';
    LogPasswordPipeline('CREATE_FLOW_INPUT', UserName, Password);

    // Create the user with NET USER (two-step to avoid 14-char LM password prompt),
    // then PowerShell fallback.
    // Step 1: create with short throwaway password (no LM prompt)
    // Step 2: set real password (password change does not trigger the LM prompt)
    // If either step fails, delete the user and fall through to PowerShell.
    OutPath := TempFile('user_create_' + SanitizeFileName(UserName) + '.log');
    NetRc := RunNetHidden('user ' + QuoteExeArg(UserName) + ' ' + QuoteExeArg(NET_USER_TEMP_PASSWORD) + ' /add /fullname:' + QuoteExeArg(UserName) + ' /expires:never');
    if NetRc = 0 then
    begin
      NetRc := RunNetHidden('user ' + QuoteExeArg(UserName) + ' ' + QuoteExeArg(Password));
      if NetRc <> 0 then
      begin
        WriteInstallerLog('WARNING: NET user password set failed for ' + UserName + ', deleting partial user');
        RunNetHidden('user ' + QuoteExeArg(UserName) + ' /delete');
      end;
    end;
    ResultCode := NetRc;

    if ResultCode <> 0 then
    begin
      UserCreatePath := 'POWERSHELL';
      WriteInstallerLog('WARNING: NET user creation failed for ' + UserName + ', falling back to PowerShell New-LocalUser path');
      PSScript :=
        'param([string]$UserName, [string]$Password, [string]$OutPath)' + #13#10 +
        '$ErrorActionPreference = ''Stop''' + #13#10 +
        'try {' + #13#10 +
        '  $outDir = [System.IO.Path]::GetDirectoryName($OutPath)' + #13#10 +
        '  if (-not [string]::IsNullOrWhiteSpace($outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }' + #13#10 +
        '  $pw = ConvertTo-SecureString -String $Password -AsPlainText -Force' + #13#10 +
        '  $existing = Get-LocalUser -Name $UserName -ErrorAction SilentlyContinue' + #13#10 +
        '  if ($null -ne $existing) {' + #13#10 +
        '    Set-LocalUser -Name $UserName -Password $pw -ErrorAction Stop' + #13#10 +
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
    NetRc := RunNetHidden('localgroup ' + QuoteExeArg(GroupAdministratorsName) + ' ' + QuoteExeArg(UserName) + ' /add');
    ResultCode := NetRc;
    if ResultCode <> 0 then
    begin
      PSScript := BuildAddGroupMemberPowerShellScript(GroupAdministratorsName, UserName, OutPath, 'ADD_ADMIN_OK');
      WriteInstallerLog('DEBUG: Fallback: Adding user to Administrators via PowerShell: Add-LocalGroupMember -Group ' + GroupAdministratorsName + ' -Member ' + UserName);
      PSParams :=
        BuildPSNamedParam('GroupName', GroupAdministratorsName) + ' ' +
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
    NetRc := RunNetHidden('localgroup ' + QuoteExeArg(GroupRDPUsersName));
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
      NetRc := RunNetHidden('localgroup ' + QuoteExeArg(GroupRDPUsersName) + ' ' + QuoteExeArg(UserName) + ' /add');
      ResultCode := NetRc;
      if ResultCode <> 0 then
      begin
        PSScript := BuildAddGroupMemberPowerShellScript(GroupRDPUsersName, UserName, OutPath, 'ADD_RDP_OK');
        WriteInstallerLog('DEBUG: Fallback: Adding user to Remote Desktop Users via PowerShell: Add-LocalGroupMember -Group ' + GroupRDPUsersName + ' -Member ' + UserName);
        PSParams :=
          BuildPSNamedParam('GroupName', GroupRDPUsersName) + ' ' +
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
    CreateRDPShortcut(UserName, Password, 'EXISTING_USER');
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
  if Assigned(L) then
  begin
    L.Caption := '✓ ' + Text;
    if IsDarkColor(WizardForm.Color) then
      L.Font.Color := RGBToColor(0, 200, 0)   // brighter green for dark backgrounds
    else
      L.Font.Color := RGBToColor(0, 128, 0);  // standard green for light backgrounds
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
  if Assigned(StepInstallTermWrap) then StepInstallTermWrap.Visible := False;
  if Assigned(StepConfigureService) then StepConfigureService.Visible := False;
  if Assigned(StepCreateUsers) then StepCreateUsers.Visible := False;
  if Assigned(StepCreateShortcuts) then StepCreateShortcuts.Visible := False;
  if Assigned(StepPreTrust) then StepPreTrust.Visible := False;
  if Assigned(StepStartSvc) then StepStartSvc.Visible := False;
  if Assigned(StepCheckRDP) then StepCheckRDP.Visible := False;
  if Assigned(StepUninstallTermWrap) then StepUninstallTermWrap.Visible := False;
  if Assigned(StepRemoveFolder) then StepRemoveFolder.Visible := False;
  if Assigned(StepShowRDPInfo) then StepShowRDPInfo.Visible := False;
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

procedure ReadShortcutSettingsFromRdpFile(const RdpPath: string);
// Reads display/audio settings from an existing .rdp file and populates the
// shortcut settings controls.  Uses PowerShell to handle both UTF-8 and
// UTF-16 LE encodings (mstsc writes UTF-16; this installer writes UTF-8).
var
  OutPath, PSScript, ScriptPath: string;
  Lines: TStringList;
  i, EqPos: Integer;
  Key, Val, Line: string;
  ScreenMode, DesktopWidth, DesktopHeight, UseMultiMon, AudioMode, RedirectClipboard: Integer;
  DisableWallpaper, AllowFontSmooth, AllowComposition: Integer;
  DisableFullWindowDrag, DisableMenuAnims, DisableThemes: Integer;
  ResIndex, ResultCode: Integer;
begin
  if not FileExists(RdpPath) then exit;

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
    '};' +
    '[System.IO.File]::WriteAllText(''' + OutPath + ''', ($out -join "`n"))';

  ScriptPath := TempFile('read_rdp_settings.ps1');
  SaveStringToFile(ScriptPath, PSScript, False);
  Exec(EXE_POWERSHELL, BuildPowerShellFileArgs(ScriptPath, '', True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  WriteInstallerLog('ReadShortcutSettingsFromRdpFile: PS exit=' + IntToStr(ResultCode) + ' path=' + RdpPath);

  if not FileExists(OutPath) then exit;

  // Defaults (match BuildShortcutSettingsBlock initial state)
  ScreenMode := 1; DesktopWidth := DEFAULT_RDP_WIDTH; DesktopHeight := DEFAULT_RDP_HEIGHT;
  UseMultiMon := 0; AudioMode := 0; RedirectClipboard := 1;
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
      else if Key = 'disable_themes'       then DisableThemes        := StrToIntDef(Val, 0);
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
end;

procedure WriteShortcutSettingsToRdpFile(const RdpPath: string);
// Reads the existing .rdp file and updates the display/audio/experience settings from the
// current shortcut settings UI controls, then writes the file back in-place.
var
  ScreenModeId, DesktopWidth, DesktopHeight, UseMultiMon, AudioMode, RedirectClipboard: Integer;
  DisableWallpaper, AllowFontSmooth, AllowComposition: Integer;
  DisableFullWindowDrag, DisableMenuAnims, DisableThemes: Integer;
  ResultCode: Integer;
  PSScript, ScriptPath: string;
begin
  // Collect values from UI controls
  GetShortcutDisplaySettings(ScreenModeId, DesktopWidth, DesktopHeight, UseMultiMon, AudioMode, RedirectClipboard);
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
    '[System.IO.File]::WriteAllLines($path, $lines)';

  ScriptPath := TempFile('update_rdp_settings.ps1');
  SaveStringToFile(ScriptPath, PSScript, False);
  Exec(EXE_POWERSHELL, BuildPowerShellFileArgs(ScriptPath, '', True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  WriteInstallerLog('WriteShortcutSettingsToRdpFile: exit=' + IntToStr(ResultCode) + ' path=' + RdpPath);
  if ResultCode = 0 then
  begin
    SignRdpFile(RdpPath);
  end;
end;

// Displays help for a setting on the Shortcut Settings page.
// Tag 1=Copy&Paste, 2=Sound, 3=Monitor settings (Window Size/Full Screen/All Monitors), 4=Performance vs Quality checkboxes
procedure ShowShortcutHelpInfo(Sender: TObject);
var
  HelpText: string;
begin
  case TButton(Sender).Tag of
    1: HelpText :=
         'Allow Copy && Paste' + #13#10#13#10 +
         'Enables clipboard sharing between the remote session and the local PC.' + #13#10#13#10 +
         'When checked, you can copy text, images, and files on one side and paste them on the other. ' +
         'When unchecked, the clipboard is isolated. Nothing can be transferred between the two sides.';
    2: HelpText :=
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
         'Disabling effects improves responsiveness and performance.' + #13#10#1310 +
         'Desktop wallpaper: shows or hides the background image in the session.' + #13#10 +
         'Smooth text (ClearType): enables font anti-aliasing for sharper text.' + #13#10 +
         'Transparent windows & effects: enables Aero glass and window animations.' + #13#10 +
         'Show contents while dragging: renders window contents as you drag them.' + #13#10 +
         'Animated menus & transitions: enables fade/slide animations on menus.' + #13#10 +
         'Visual themes: enables Windows visual styling (disabling gives a classic look).';
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

  // Current file context (used by Edit Shortcut flow)
  lblShortcutEditingFile := TLabel.Create(ParentSurface);
  lblShortcutEditingFile.Parent := ParentSurface;
  lblShortcutEditingFile.Left := ScaleX(10);
  lblShortcutEditingFile.Top := StartTop;
  lblShortcutEditingFile.Caption := 'Editing:';
  lblShortcutEditingFile.Font.Style := [fsBold];
  lblShortcutEditingFile.AutoSize := True;

  // Section separator label
  lblShortcutSection := TLabel.Create(ParentSurface);
  lblShortcutSection.Parent := ParentSurface;
  lblShortcutSection.Left := ScaleX(10);
  lblShortcutSection.Top := lblShortcutEditingFile.Top + ScaleY(22);
  lblShortcutSection.Caption := 'Basic Shortcut Settings';
  lblShortcutSection.Font.Style := [fsBold];
  lblShortcutSection.AutoSize := True;

  // Row 1 — Copy & Paste
  chkCopyPaste := TCheckBox.Create(ParentSurface);
  chkCopyPaste.Parent := ParentSurface;
  chkCopyPaste.Left := ScaleX(10);
  chkCopyPaste.Top := lblShortcutSection.Top + ScaleY(24);
  chkCopyPaste.Width := ScaleX(150);
  chkCopyPaste.Caption := 'Allow Copy && Paste';
  chkCopyPaste.Checked := True;
  MakeShortcutHelpButton(ParentSurface, chkCopyPaste.Top, 1);

  // Row 2 — Sound (own row)
  chkSound := TCheckBox.Create(ParentSurface);
  chkSound.Parent := ParentSurface;
  chkSound.Left := ScaleX(10);
  chkSound.Top := chkCopyPaste.Top + ScaleY(24);
  chkSound.Width := ScaleX(150);
  chkSound.Caption := 'Allow Sound';
  chkSound.Checked := True;
  MakeShortcutHelpButton(ParentSurface, chkSound.Top, 2);

  // Row 3 — Screen Size label
  lblScreenSize := TLabel.Create(ParentSurface);
  lblScreenSize.Parent := ParentSurface;
  lblScreenSize.Left := ScaleX(10);
  lblScreenSize.Top := chkSound.Top + ScaleY(26);
  lblScreenSize.Caption := 'Window Size:';
  lblScreenSize.AutoSize := True;

  // Row 3 — Resolution drop-down
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

  // Row 3 — Full Screen checkbox
  chkFullScreen := TCheckBox.Create(ParentSurface);
  chkFullScreen.Parent := ParentSurface;
  chkFullScreen.Left := cboResolution.Left + cboResolution.Width + ScaleX(10);
  chkFullScreen.Top := lblScreenSize.Top - ScaleY(1);
  chkFullScreen.Width := ScaleX(80);
  chkFullScreen.Caption := 'Full Screen';
  chkFullScreen.Checked := False;
  chkFullScreen.OnClick := @OnFullScreenClick;
  cboResolution.Enabled := True;

  // Row 3 — Use All Monitors
  chkUseAllMonitors := TCheckBox.Create(ParentSurface);
  chkUseAllMonitors.Parent := ParentSurface;
  chkUseAllMonitors.Left := chkFullScreen.Left + chkFullScreen.Width + ScaleX(20);
  chkUseAllMonitors.Top := lblScreenSize.Top - ScaleY(1);
  chkUseAllMonitors.Width := ScaleX(110);
  chkUseAllMonitors.Caption := 'Use All Monitors';
  chkUseAllMonitors.Checked := False;
  chkUseAllMonitors.OnClick := @OnUseAllMonitorsClick;
  MakeShortcutHelpButton(ParentSurface, lblScreenSize.Top - ScaleY(1), 3);

  // Row 3b — Custom resolution W/H inputs (hidden until "Custom" is selected)
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

  // Row 4 — Performance section header (offset by extra row to clear Custom resolution inputs)
  TmpLabel := TLabel.Create(ParentSurface);
  TmpLabel.Parent := ParentSurface;
  TmpLabel.Left := ScaleX(10);
  TmpLabel.Top := lblScreenSize.Top + ScaleY(52);
  TmpLabel.Caption := 'Quality vs Performance - (Unchecked = better performance):';
  TmpLabel.Font.Style := [fsBold];
  TmpLabel.AutoSize := True;
  MakeShortcutHelpButton(ParentSurface, TmpLabel.Top, 4);

  // Row 5 — Experience checkboxes (2 columns, 3 rows)
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

  // Row 5 — Tips for editing more settings
  lblShortcutTips := TLabel.Create(ParentSurface);
  lblShortcutTips.Parent := ParentSurface;
  lblShortcutTips.Left := ScaleX(10);
  lblShortcutTips.Top := chkExpMenuAnim.Top + ScaleY(28);
  lblShortcutTips.Caption := 'For more settings, run this app again and choose "Edit existing shortcut settings"';
  lblShortcutTips.Font.Style := [fsItalic];
  lblShortcutTips.AutoSize := True;

  // Row 6 — Multi-user note (shown when 2+ shortcuts will be created, hidden otherwise)
  lblMultiShortcutEditingNote := TLabel.Create(ParentSurface);
  lblMultiShortcutEditingNote.Parent := ParentSurface;
  lblMultiShortcutEditingNote.Left := ScaleX(10);
  lblMultiShortcutEditingNote.Top := lblShortcutTips.Top + ScaleY(26);
  lblMultiShortcutEditingNote.Font.Style := [fsBold];
  lblMultiShortcutEditingNote.AutoSize := True;
  lblMultiShortcutEditingNote.Visible := False;

  // Row 4 — "Open advanced shortcut options" checkbox (shown only in Edit Shortcuts mode)
  chkShowMoreShortcutOptions := TCheckBox.Create(ParentSurface);
  chkShowMoreShortcutOptions.Parent := ParentSurface;
  chkShowMoreShortcutOptions.Left := ScaleX(260);
  chkShowMoreShortcutOptions.Top := lblShortcutTips.Top + ScaleY(26);
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
  // Initialize installer log file
  InitInstallerLog;
  WriteInstallerLog('BUILD_FINGERPRINT=' + BUILD_FINGERPRINT);
  WriteInstallerLog('PRESERVE_USER_CREATE_DEBUG_LOGS=' + BoolToStr(PRESERVE_USER_CREATE_DEBUG_LOGS <> 0));
  WriteInstallerLog('PASSWORD_PIPELINE_DIAG=' + BoolToStr(PASSWORD_PIPELINE_DIAG <> 0));

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
  if IsTermWrapInstalled() then
  begin
    chkInstallTermWrap.Caption := 'Install TermWrap (Already installed. Selecting this will re-install it)';
    chkInstallTermWrap.Checked := False;
  end
  else
  begin
    chkInstallTermWrap.Caption := 'Install TermWrap';
    chkInstallTermWrap.Checked := True;
  end;

  chkCreateRdpShortcuts := TCheckBox.Create(Page_InstallOptions);
  chkCreateRdpShortcuts.Parent := Page_InstallOptions.Surface;
  chkCreateRdpShortcuts.Left := ScaleX(30);
  chkCreateRdpShortcuts.Top := ScaleY(60);
  chkCreateRdpShortcuts.Width := ScaleX(380);
  chkCreateRdpShortcuts.Caption := 'Create RDP shortcuts';
  chkCreateRdpShortcuts.Checked := True;
  chkCreateRdpShortcuts.OnClick := @OnCreateRdpShortcutsClick;

  CreateRdpShortcutsGroup := TPanel.Create(Page_InstallOptions);
  CreateRdpShortcutsGroup.Parent := Page_InstallOptions.Surface;
  CreateRdpShortcutsGroup.Left := ScaleX(40);
  CreateRdpShortcutsGroup.Top := ScaleY(84);
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

  // Edit Shortcut radio placed halfway between Install and Uninstall

  // Set initial enabled state based on Create RDP shortcuts checkbox
  rbCreateUsers.Enabled := chkCreateRdpShortcuts.Checked;
  rbUseExistingUsers.Enabled := chkCreateRdpShortcuts.Checked;



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

  rbUninstall := TRadioButton.Create(Page_InstallOptions);
  rbUninstall.Parent := Page_InstallOptions.Surface;
  rbUninstall.Left := ScaleX(10);
  rbUninstall.Top := radioTopBase + radioSpacing * 3;
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
  lblBullet5.Caption := '• Special thanks to Bee Swarm Sim communities: ';
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
  lblBullet4.Caption := '• Assembled by cpdx4. Project Home: ';
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
    'Configure the settings for your RDP desktop shortcuts. If unsure, just click [Next]'
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
    'Select one Desktop .rdp shortcut to edit.'
  );
  
  
  // Initialize lists for tracking
  LocalUsersList := TStringList.Create;        // Will be populated when Create Shortcuts page is shown
  LocalUserDisplayList := TStringList.Create;  // Parallel display labels (email for online accounts)
  DesktopRdpFiles := TStringList.Create;
  SetLength(UserCheckBoxes, 0);
  SetLength(UserPasswordEdits, 0);
  SetLength(UserPasswordStatus, 0);
  SetLength(ShortcutRadioButtons, 0);
  CurrentShortcutPage := 0;
  ShortcutsPerPage := 8;
  SelectedShortcutIndex := -1;
  SelectedShortcutPath := '';
  EditShortcutControlsBuilt := False;
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
begin
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

    if rbUninstall.Checked then
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
      for i := 0 to Length(ShortcutRadioButtons) - 1 do
        if Assigned(ShortcutRadioButtons[i]) then ShortcutRadioButtons[i].Visible := False;
      EditShortcutControlsBuilt := False;
      SetLength(ShortcutRadioButtons, 0);
      SelectedShortcutIndex := -1;
      SelectedShortcutPath := '';
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
  end
  else if CurPageID = EditShortcutPage.ID then
  begin
    if DesktopRdpFiles.Count = 0 then
    begin
      MsgBox('No .rdp shortcuts were found on your Desktop.', mbError, MB_OK);
      Result := False;
      exit;
    end;

    if (SelectedShortcutIndex < 0) or (SelectedShortcutIndex >= DesktopRdpFiles.Count) then
    begin
      MsgBox('Select a shortcut to edit before continuing.', mbError, MB_OK);
      Result := False;
      exit;
    end;

    // Record the selected shortcut path; do not launch mstsc here.
    // The editor will be launched as part of the Installing step to keep
    // all external launches inside the install flow.
    SelectedShortcutPath := DesktopRdpFiles[SelectedShortcutIndex];
  end
  else if CurPageID = Page_ShortcutSettings.ID then
  begin
    if SelectedInstallMode = installModeEditShortcuts then
    begin
      // Remember whether the user wants to open the advanced mstsc editor
      DoShowMstscEdit := Assigned(chkShowMoreShortcutOptions) and chkShowMoreShortcutOptions.Checked;
      // Apply the simple settings from this page to the selected .rdp file now,
      // before the installing step (mstsc /edit can then further edit it)
      if SelectedShortcutPath <> '' then
        WriteShortcutSettingsToRdpFile(SelectedShortcutPath);
    end;
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
  MstscPath: string;
  i: Integer;
  UserInfo: string;
  UserName: string;
  Password: string;
begin
  if CurStep = ssInstall then
  begin
    LogSectionHeader('STEP TRANSITION: ssInstall');
    LogKeyValue('SelectedInstallMode', IntToStr(SelectedInstallMode));
    LogKeyValue('DoInstallTermWrap', BoolToStr(DoInstallTermWrap));
    LogKeyValue('DoCreateRdpShortcuts', BoolToStr(DoCreateRdpShortcuts));
    LogKeyValue('CreateUserMode', IntToStr(CreateUserMode));
    // Hide cancel button during installation to prevent confusion
    WizardForm.CancelButton.Visible := False;
    
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
        AddStepPendingLabel(StepInstallTermWrap, TXT_InstallTermWrap);
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
      AddStepPendingLabel(StepCreateShortcuts, 'Open selected .rdp in editor');
    end;

    if (SelectedInstallMode = installModeInstall) and DoInstallTermWrap then
      CheckAndInstallMSTSC;

    // Handle uninstall cleanup
    if SelectedInstallMode = installModeUninstall then
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
      WizardForm.StatusLabel.Caption := 'Preparing Create Shortcuts...';
      WizardForm.ProgressGauge.Style := npbstMarquee;
    end
    // Tune performance: deferred to ssPostInstall
    else if SelectedInstallMode = installModeShowRDPInfo then
    begin
      WizardForm.StatusLabel.Caption := 'Preparing to apply Group Policy settings...';
      WizardForm.ProgressGauge.Style := npbstMarquee;
    end
    else if SelectedInstallMode = installModeEditShortcuts then
    begin
      WizardForm.StatusLabel.Caption := 'Preparing shortcut editor completion...';
      WizardForm.ProgressGauge.Style := npbstMarquee;
      SetStepDone(StepCreateShortcuts, 'Open selected .rdp in editor');
    end
    // Only stop TermService when installing TermWrap
    else if (SelectedInstallMode = installModeInstall) and DoInstallTermWrap then
    begin
      WizardForm.StatusLabel.Caption := 'Preparing installation...';
      WizardForm.ProgressGauge.Style := npbstMarquee;
      
      // Now that UI is visible, safely stop the service (executes first, displays first)
      SetStepInProgress(StepStopSvc, TXT_StopSvc);
      WizardForm.StatusLabel.Caption := 'Stopping Remote Desktop Services...';
      Log('[CurStepChanged-ssInstall] Stopping TermService for Install TermWrap');
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
    LogSectionHeader('STEP TRANSITION: ssPostInstall');
    // Handle uninstall completion
    if SelectedInstallMode = installModeUninstall then
    begin
      WizardForm.StatusLabel.Caption := 'Uninstallation complete! TermWrap has been removed.';
    end
    // Show RDP Info: display-only page, no settings to apply
    else if SelectedInstallMode = installModeShowRDPInfo then
    begin
      SetStepInProgress(StepShowRDPInfo, TXT_ShowRDPInfo);
      SetStepDone(StepShowRDPInfo, TXT_ShowRDPInfo);
      WizardForm.StatusLabel.Caption := 'Done.';
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

      WizardForm.StatusLabel.Caption := 'System changes applied.';
    end
    // Edit System-wide settings fallback (no changes queued): create shortcuts for existing users
    else if SelectedInstallMode = installModeEditSystemwideSettings then
    begin
      SetStepInProgress(StepCreateShortcuts, TXT_CreateShortcuts);
      WizardForm.StatusLabel.Caption := 'Creating RDP shortcuts...';
      CreateShortcutsForExistingUsers;
      // Create Shortcuts path completed
      SetStepDone(StepCreateShortcuts, TXT_CreateShortcuts);
      SetStepInProgress(StepPreTrust, TXT_PreTrust);
      WizardForm.StatusLabel.Caption := 'Pre-trusting Remote Desktop certificate...';
      PreTrustRDPCertCurrentUser;
      SetStepDone(StepPreTrust, TXT_PreTrust);
      ClearPasswordsFromMemory;
      WizardForm.StatusLabel.Caption := 'Create Shortcuts executed.';
    end
    else if SelectedInstallMode = installModeEditShortcuts then
    begin
      // Apply settings to the .rdp file and optionally open mstsc /edit
      SetStepInProgress(StepCreateShortcuts, 'Open selected .rdp in editor');
      if DoShowMstscEdit then
      begin
        WizardForm.StatusLabel.Caption := 'Opening selected .rdp in editor...';
        if SelectedShortcutPath = '' then
        begin
          WriteInstallerLog('Edit Shortcut: no SelectedShortcutPath set');
        end
        else
        begin
          // Try to launch mstsc to edit the selected shortcut
          MstscPath := GetMstscPath;
          if (MstscPath = '') or (not Exec(MstscPath, '/edit "' + SelectedShortcutPath + '"', '', SW_SHOW, ewNoWait, ResultCode)) then
          begin
            MsgBox('Failed to launch Remote Desktop editor. Verify mstsc is available and try again.', mbError, MB_OK);
            WriteInstallerLog('Edit Shortcut: Exec(mstsc) failed exit=' + IntToStr(ResultCode));
          end;
        end;
      end
      else
      begin
        WizardForm.StatusLabel.Caption := 'Shortcut settings applied.';
        WriteInstallerLog('Edit Shortcut: mstsc editor skipped by user choice');
      end;
      SetStepDone(StepCreateShortcuts, 'Open selected .rdp in editor');
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
        WizardForm.StatusLabel.Caption := 'Downloading VC++ Redistributable from Microsoft...';
        WizardForm.ProgressGauge.Style := npbstMarquee;
        
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
        
        WizardForm.StatusLabel.Caption := 'Installing VC++ Redistributable (this may take a minute)...';
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
        WizardForm.StatusLabel.Caption := 'VC++ Redistributable already installed, skipping...';
      end;
      // VC++ ensured (installed or skipped)
      SetStepDone(StepEnsureVC, TXT_EnsureVC);
      
      // Install TermWrap
      SetStepInProgress(StepInstallTermWrap, TXT_InstallTermWrap);
      WizardForm.StatusLabel.Caption := 'Installing TermWrap...';
      // TermWrap files are bundled and copied earlier; no external installer to run.
      Sleep(SLEEP_SHORT);
      SetStepDone(StepInstallTermWrap, TXT_InstallTermWrap);
      
      SetStepInProgress(StepConfigureService, TXT_ConfigureService);
      WizardForm.StatusLabel.Caption := 'Configuring TermWrap service...';
      
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

      // Hide most security warnings by default for TermWrap installs
      if RegWriteDWordValue(HKLM, REG_TS_POLICIES + '\\Client', 'RedirectionWarningDialogVersion', 1) then
        WriteInstallerLog('Registry: Set RedirectionWarningDialogVersion=1 (Hide security warnings)')
      else
        WriteInstallerLog('Registry: FAILED to set RedirectionWarningDialogVersion');

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
    
    // Pre-trust for current user in all Install sub-flows (AFTER service starts if applicable)
    if SelectedInstallMode = installModeInstall then
    begin
      SetStepInProgress(StepPreTrust, TXT_PreTrust);
      WizardForm.StatusLabel.Caption := 'Pre-trusting Remote Desktop certificate...';
      PreTrustRDPCertCurrentUser;
      // Pre-trust is optional - don't fail install if cert doesn't exist yet
      SetStepDone(StepPreTrust, TXT_PreTrust);
    end;

    // Verify RDP is listening (only when TermWrap was installed)
    if (SelectedInstallMode = installModeInstall) and DoInstallTermWrap then
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

procedure CurPageChanged(CurPageID: Integer);
var
  i: Integer;
  CompletionText: string;
  Entry: string;
  UserName: string;
  Password: string;
  NowTick: Cardinal;
  IsDuplicatePageEvent: Boolean;
  DeltaMs: Cardinal;
  DisplayVersion: string;
  BuildNumberStr: string;
  UBRVal: Cardinal;
  ServiceStatus: string;
  TermsrvVer: string;
  ServiceDllPath: string;
  TermWrapVer: string;
  PortNumber: Cardinal;
begin
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

  // Configure shortcut settings page based on which flow is entering it
  if CurPageID = Page_ShortcutSettings.ID then
  begin
    WizardForm.NextButton.Caption := SetupMessage(msgButtonNext);
    WizardForm.ActiveControl := WizardForm.NextButton;  // prevent cboResolution from receiving initial focus

    if SelectedInstallMode = installModeEditShortcuts then
    begin
      if Assigned(lblShortcutEditingFile) then
      begin
        if SelectedShortcutPath <> '' then
          lblShortcutEditingFile.Caption := 'Editing:  ' + ExtractFileName(SelectedShortcutPath)
        else
          lblShortcutEditingFile.Caption := 'Editing:';
      end;
      // EditShortcuts path: hide multi-note and tips, show "more options" checkbox
      if Assigned(lblMultiShortcutEditingNote) then lblMultiShortcutEditingNote.Visible := False;
      if Assigned(lblShortcutTips) then lblShortcutTips.Visible := False;
      if Assigned(chkShowMoreShortcutOptions) then chkShowMoreShortcutOptions.Visible := True;
      // Pre-populate controls with the current settings from the selected .rdp file
      if SelectedShortcutPath <> '' then
        ReadShortcutSettingsFromRdpFile(SelectedShortcutPath);
    end
    else if (SelectedInstallMode = installModeInstall) and (CreateUserMode = createUserModeNew) then
    begin
      if Assigned(lblShortcutEditingFile) then
        lblShortcutEditingFile.Caption := 'Editing:  New shortcut(s)';
      // CreateUsers path: show tips, show multi-note only when 2+ users queued
      if Assigned(chkShowMoreShortcutOptions) then chkShowMoreShortcutOptions.Visible := False;
      if Assigned(lblShortcutTips) then lblShortcutTips.Visible := True;
      if Assigned(lblMultiShortcutEditingNote) then
      begin
        if UsersList.Count > 1 then
        begin
          lblMultiShortcutEditingNote.Caption := 'These settings will be applied to each shortcut';
          lblMultiShortcutEditingNote.Visible := True;
        end
        else
          lblMultiShortcutEditingNote.Visible := False;
      end;
    end
    else
    begin
      if Assigned(lblShortcutEditingFile) then
        lblShortcutEditingFile.Caption := 'Editing:  Selected shortcut(s)';
      // ExistingUsers path: hide tips, show multi-note only when 2+ shortcuts selected
      if Assigned(chkShowMoreShortcutOptions) then chkShowMoreShortcutOptions.Visible := False;
      if Assigned(lblShortcutTips) then lblShortcutTips.Visible := False;
      if Assigned(lblMultiShortcutEditingNote) then
      begin
        if ShortcutsList.Count > 1 then
        begin
          lblMultiShortcutEditingNote.Caption := 'These settings will be applied to each shortcut';
          lblMultiShortcutEditingNote.Visible := True;
        end
        else
          lblMultiShortcutEditingNote.Visible := False;
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

  // Populate Show RDP Info page when shown
  if Assigned(Page_ShowRDPInfo) and (CurPageID = Page_ShowRDPInfo.ID) then
  begin
    // Reset displayed values to indicate loading when page is shown again (e.g. after Back/Next)
    if Assigned(lblWinVer) then lblWinVer.Caption := '--';
    if Assigned(lblRDPService) then lblRDPService.Caption := '--';
    if Assigned(lblWinRDPVer) then lblWinRDPVer.Caption := '--';
    if Assigned(lblWrapperVer) then lblWrapperVer.Caption := '--';

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

    ServiceDllPath := ExpandConstant('{commonpf64}\RDPWrapKit\TermWrap.dll');
    if FileExists(ServiceDllPath) then
    begin
      TermWrapVer := GetPSOutput('(Get-Item -Path ''' + ServiceDllPath + ''').VersionInfo.FileVersion');
      if TermWrapVer = '' then lblWrapperVer.Caption := 'TermWrap (unknown version)'
      else lblWrapperVer.Caption := 'TermWrap ' + TermWrapVer;
    end
    else
    begin
      ServiceDllPath := ExpandConstant('{commonpf64}\RDP Wrapper\rdpwrap.dll');
      if FileExists(ServiceDllPath) then
      begin
        TermWrapVer := GetPSOutput('(Get-Item -Path ''' + ServiceDllPath + ''').VersionInfo.FileVersion');
        if TermWrapVer = '' then lblWrapperVer.Caption := 'RDPWrap (unknown version)'
        else lblWrapperVer.Caption := 'RDPWrap ' + TermWrapVer;
      end
      else
        lblWrapperVer.Caption := 'None (Windows default)';
    end;

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
    else if SelectedInstallMode = installModeEditShortcuts then
    begin
      if DoShowMstscEdit then
      begin
        WizardForm.FinishedHeadingLabel.Caption := 'Shortcut Editor Complete';
        CompletionText :=
          'The shortcut was opened in the Remote Desktop Connection app. Make your changes there.' + #13#10#13#10 +
          'Important: Always click [Save] on the General tab.';
        WriteInstallerLog('CurPageChanged: Showing shortcut editor completion message');

        if Assigned(FinishedExampleImage) then
        begin
          FinishedExampleImage.Top := FinishedText.Top + ScaleY(100);
          FinishedExampleImage.Width := ScaleX(142);
          FinishedExampleImage.Height := ScaleY(150);
          FinishedExampleImage.Stretch := False;
          // Load the rdp_edit_save.bmp image from temp
          try
            ExtractTemporaryFile(FILE_RDPEDITSAVE_BMP);
            FinishedExampleImage.Bitmap.LoadFromFile(ExpandConstant(TEMP_RDPEDITSAVE_BMP));
          except
            // If image fails to load, hide the control
            FinishedExampleImage.Visible := False;
          end;
          FinishedExampleImage.Visible := True;
          
        end;
      end
      else
      begin
        WizardForm.FinishedHeadingLabel.Caption := 'Shortcut Settings Updated';
        CompletionText := 'Your shortcut settings have been saved to the .rdp file.';
        WriteInstallerLog('CurPageChanged: Showing shortcut settings saved message (no mstsc edit)');
        if Assigned(FinishedExampleImage) then FinishedExampleImage.Visible := False;
      end;
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
          CompletionText := CompletionText + 'You can now open RDP connections using the created shortcuts.';
      end;
    end;
    
    // Add note about Windows security warning when connecting to remote PCs
    CompletionText := CompletionText + #13#10#13#10 +
      'Windows may show a security warning when you connect to remote PCs. This is normal. RDPWrapKit is local so the connection never leaves your network.';

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
begin
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
    Log('DEBUG: mstsc.exe found.');
    LogKeyValue('mstsc path', MstscPath);
    SetStepDone(StepCheckMSTSC, TXT_CheckMSTSC);
    SetStepDone(StepInstallMSTSC, TXT_InstallMSTSC); // skipped
  end
  else
  begin
    Log('DEBUG: mstsc.exe missing. Initiating installation.');
    LogKeyValue('Download URL', URL_RDP_INSTALLER);
    SetStepDone(StepCheckMSTSC, TXT_CheckMSTSC);
    SetStepInProgress(StepInstallMSTSC, TXT_InstallMSTSC);
    InstallerPath := TempFile('mstsc_installer.exe');
    LogKeyValue('Installer temp path', InstallerPath);
    Exec(EXE_POWERSHELL, BuildPowerShellArgs('$out = ''' + InstallerPath + '''; Invoke-WebRequest -Uri ''' + URL_RDP_INSTALLER + ''' -OutFile $out -UseBasicParsing', True), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    LogKeyValue('Download exit code', IntToStr(ResultCode));
    LogKeyValue('Installer file exists after download', BoolToStr(FileExists(InstallerPath)));
    if (ResultCode = 0) and FileExists(InstallerPath) and IsSignedByMicrosoftCorporation(InstallerPath) then
    begin
      Exec(InstallerPath, '', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      LogKeyValue('Installer execution exit code', IntToStr(ResultCode));
      if ResultCode = 0 then
      begin
        Log('DEBUG: Remote Desktop Connection installed successfully.');
        SetStepDone(StepInstallMSTSC, TXT_InstallMSTSC);
      end
      else
      begin
        Log('ERROR: mstsc installer execution failed. Exit code: ' + IntToStr(ResultCode));
        PromptManualDownload('Remote Desktop Connection (mstsc)', URL_RDP_INSTALLER, 'Installer execution failed (exit code ' + IntToStr(ResultCode) + ')');
        SetStepDone(StepInstallMSTSC, TXT_InstallMSTSC);
      end;
    end
    else
    begin
      Log('ERROR: Failed to download or validate Remote Desktop Connection installer. Exit code: ' + IntToStr(ResultCode));
      PromptManualDownload('Remote Desktop Connection (mstsc)', URL_RDP_INSTALLER, 'Download or signature validation failed (exit code ' + IntToStr(ResultCode) + ')');
      SetStepDone(StepInstallMSTSC, TXT_InstallMSTSC); // mark as done even on failure?
      // Perhaps leave it pending or something, but for now, done.
    end;
    if FileExists(InstallerPath) then
      DeleteFile(InstallerPath);
  end;
end;
