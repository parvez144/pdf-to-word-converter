; PDF to Word Converter - Inno Setup Script
; Created for professional installation

#define MyAppName "PDF to Word Converter"
#define MyAppVersion "1.0"
#define MyAppPublisher "Shahriar Parvez Khan"
#define MyAppURL "https://github.com/parvez144/pdf-to-word-converter"
#define MyAppExeName "PDF to Word Converter.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
LicenseFile=LICENSE.txt
; Uncomment the following line to run in debug mode.
;Debugging=yes
InfoBeforeFile=INFO_BEFORE.txt
InfoAfterFile=INFO_AFTER.txt
; Remove the following line to run in administrative install mode (install for all users.)
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=Output
OutputBaseFilename=PDF_to_Word_Converter_Setup
SetupIconFile=spk.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; WizardImageFile=spk.ico
; WizardSmallImageFile=spk.ico
; Alternative image syntax examples:
; WizardImageFile=wizard.bmp        ; Bitmap format (164x314 pixels)
; WizardImageFile=wizard.png        ; PNG format (164x314 pixels)
; WizardSmallImageFile=small.bmp    ; Small bitmap (55x55 pixels)
; WizardSmallImageFile=small.png    ; Small PNG (55x55 pixels)
; SetupIconFile=app.ico            ; Different icon file
DisableProgramGroupPage=no
DisableDirPage=no
DisableReadyPage=no
DisableFinishedPage=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "spk.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.txt"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\spk.ico"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"; IconFilename: "{app}\spk.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\spk.ico"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\spk.ico"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
  // Add any custom initialization code here
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    // Add any post-installation code here
  end;
end;

[Registry]
; Add file association for PDF files (optional)
Root: HKCU; Subkey: "Software\Classes\.pdf\OpenWithProgids"; ValueType: string; ValueName: "PDFToWordConverter.pdf"; ValueData: ""; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Classes\PDFToWordConverter.pdf"; ValueType: string; ValueName: ""; ValueData: "PDF Document"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\PDFToWordConverter.pdf\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\spk.ico,0"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\PDFToWordConverter.pdf\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Flags: uninsdeletekey

[UninstallDelete]
Type: files; Name: "{app}\spk.ico"
Type: files; Name: "{app}\LICENSE.txt"
Type: files; Name: "{app}\README.txt" 
