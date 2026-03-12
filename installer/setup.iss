; setup.iss - Inno Setup 6 installer script for Scorecard Creator
;
; Prerequisites:
;   1. Build the PyInstaller bundle first:  pyinstaller scorecard_creator.spec
;   2. Place app_icon.ico in this installer/ folder
;   3. (Optional) Place MicrosoftEdgeWebview2Setup.exe here for WebView2 bundling
;
; Build: Open in Inno Setup 6 IDE, press Ctrl+F9
; Output: installer/output/ScorecardCreator_Setup_v1.0.0.exe

#define AppName "Scorecard Creator"
#define AppVersion "1.0.0"
#define AppPublisher "City of Cape Town"
#define AppExeName "ScorecardCreator.exe"
#define SourceDir "..\dist\ScorecardCreator"

[Setup]
; IMPORTANT: Generate a new GUID for each new application.
; In PowerShell: [System.Guid]::NewGuid().ToString().ToUpper()
AppId={{B7C4D2E8-3F91-4A05-8B62-D1E456789012}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=https://www.capetown.gov.za
AppSupportURL=https://www.capetown.gov.za
AppUpdatesURL=https://www.capetown.gov.za
DefaultDirName={autopf}\ScorecardCreator
DefaultGroupName={#AppName}
AllowNoIcons=no
OutputDir=.\output
OutputBaseFilename=ScorecardCreator_Setup_v{#AppVersion}
SetupIconFile=app_icon.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
MinVersion=10.0
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName}
RestartIfNeededByRun=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Main application files from PyInstaller dist folder
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Optional: WebView2 bootstrapper (download from Microsoft and place here)
; Source: "MicrosoftEdgeWebview2Setup.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Icons]
; Start Menu
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
; Desktop (optional - user must tick the checkbox)
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
; Optional: Install WebView2 runtime if not present
; Filename: "{tmp}\MicrosoftEdgeWebview2Setup.exe"; Parameters: "/silent /install"; \
;   StatusMsg: "Installing Microsoft Edge WebView2..."; \
;   Check: WebView2NotInstalled; Flags: skipifdoesntexist

; Offer to launch the app immediately after installation
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; User template data in AppData is preserved on uninstall by default.
; Uncomment the line below ONLY if you want a clean uninstall that removes user data:
; Type: filesandordirs; Name: "{userappdata}\ScorecardCreator"

[Code]
function WebView2NotInstalled(): Boolean;
var
  RegValue: String;
begin
  // Check for WebView2 runtime in the registry
  Result := not RegQueryStringValue(
    HKEY_LOCAL_MACHINE,
    'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{56EB18F8-8008-4CBD-B6D2-8C97FE7E9062}',
    'pv',
    RegValue
  );
end;
