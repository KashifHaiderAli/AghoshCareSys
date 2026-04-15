; AghoshCareSys Inno Setup Script
; Created for Alkhidmat Khawateen Trust Pakistan

#define MyAppName "AghoshCareSys"
#define MyAppVersion "1.0"
#define MyAppPublisher "Alkhidmat Khawateen Trust Pakistan"
#define MyAppExeName "AghoshCareSys.exe"
#define MyAppIcon "app.ico"

[Setup]
; Basic Application Information
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://www.alkhidmat.org
AppSupportURL=https://www.alkhidmat.org
AppUpdatesURL=https://www.alkhidmat.org
DefaultDirName=D:\Aghosh
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=Output
OutputBaseFilename=AghoshCareSys_Setup
SetupIconFile={#MyAppIcon}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
DisableDirPage=no
DisableProgramGroupPage=no
UninstallDisplayIcon={app}\{#MyAppExeName}

; Wizard Images (optional - remove if you don't have these)
;WizardImageFile=installer_image.bmp
;WizardSmallImageFile=installer_small.bmp

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
; Main Executable
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; Database File
Source: "Aghosh.db"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist uninsneveruninstall

; Logo Folder
Source: "logo\*"; DestDir: "{app}\logo"; Flags: ignoreversion recursesubdirs createallsubdirs

; Photo Folder (placeholder and existing photos)
Source: "Photo\*"; DestDir: "{app}\Photo"; Flags: ignoreversion recursesubdirs createallsubdirs onlyifdoesntexist uninsneveruninstall

; Application Icon
Source: "{#MyAppIcon}"; DestDir: "{app}"; Flags: ignoreversion

; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
; Start Menu Icons
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppIcon}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

; Desktop Icon (if user selected it)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppIcon}"; Tasks: desktopicon

; Quick Launch Icon (if user selected it)
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
; Run the application after installation (optional)
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Clean up empty directories on uninstall
Type: filesandordirs; Name: "{app}\logo"
; Photo folder is NOT deleted (uninsneveruninstall flag protects user data)

[Code]
// Custom message during installation
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    MsgBox('Installation completed successfully!' + #13#10 + 
           'AghoshCareSys has been installed to: D:\Aghosh' + #13#10#13#10 +
           'IMPORTANT: Your database and photos are preserved during updates.',
           mbInformation, MB_OK);
  end;
end;

// Warning before uninstall
function InitializeUninstall(): Boolean;
begin
  Result := True;
  if MsgBox('This will remove AghoshCareSys from your computer.' + #13#10#13#10 +
            'NOTE: Your database (Aghosh.db) and Photo folder will NOT be deleted.' + #13#10 +
            'You can safely reinstall later without losing data.' + #13#10#13#10 +
            'Do you want to continue?',
            mbConfirmation, MB_YESNO) = IDNO then
  begin
    Result := False;
  end;
end;