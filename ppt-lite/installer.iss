; Inno Setup script: ppt-lite installer
; Usage: iscc installer.iss  ->  output/ppt-lite-install.exe
; NOTE: keep this file ASCII-only (Inno Setup 6.x .iss parsing is picky about non-ASCII comments)
#define AppName "ppt-lite"
#define AppVersion "1.0.0"
#define AppExe "ppt-lite.exe"

[Setup]
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=ppt-lite
DefaultDirName={autopf}\ppt-lite
DefaultGroupName={#AppName}
OutputDir=output
OutputBaseFilename=ppt-lite-install
Compression=lzma2/ultra64
SolidCompression=yes
PrivilegesRequiredOverridesAllowed=dialog
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#AppExe}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional tasks:"

[Files]
Source: "dist\ppt-lite\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\ppt-lite"; Filename: "{app}\{#AppExe}"; Comment: "Start ppt-lite (open http://127.0.0.1:8765)"
Name: "{autodesktop}\ppt-lite"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon
Name: "{group}\Uninstall ppt-lite"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#AppExe}"; Description: "Launch ppt-lite now"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{localappdata}\ppt-lite\data\tmp"
Type: filesandordirs; Name: "{localappdata}\ppt-lite\data\trash"

[Code]
var
  LOPage: TInputDirWizardPage;

procedure InitializeWizard();
begin
  LOPage := CreateInputDirPage(wpSelectDir,
    'LibreOffice Location (Optional)',
    'ppt-lite uses LibreOffice for page previews and legacy .ppt conversion.',
    'If LibreOffice is installed, select its "program" folder (the one containing soffice.exe), e.g. ' +
    ExpandConstant('{autopf}') + '\LibreOffice\program' + #13#10#13#10 +
    'Leave it empty to skip. You can also configure it later with the PPT_LITE_SOFFICE environment variable.' + #13#10 +
    'Without LibreOffice everything still works, but previews are degraded (no rendered images).',
    False, '');
  LOPage.Add('LibreOffice program folder (optional):');
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  P: String;
begin
  Result := True;
  if CurPageID = LOPage.ID then begin
    P := Trim(LOPage.Values[0]);
    if (P <> '') and not FileExists(AddBackslash(P) + 'soffice.exe') then begin
      MsgBox('soffice.exe was not found in this folder. Please pick the LibreOffice "program" folder, or clear the field to skip.', mbError, MB_OK);
      Result := False;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  P, Fn: String;
begin
  if CurStep = ssPostInstall then begin
    P := Trim(LOPage.Values[0]);
    if P <> '' then begin
      Fn := ExpandConstant('{localappdata}\ppt-lite\soffice.txt');
      ForceDirectories(ExtractFileDir(Fn));
      SaveStringToFile(Fn, AddBackslash(P) + 'soffice.exe', False);
    end;
  end;
end;
