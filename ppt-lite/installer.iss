; Inno Setup 脚本：ppt-lite 安装包
; 用法：iscc installer.iss  → 产物 output/ppt-lite-install.exe
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
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务:"

[Files]
Source: "dist\ppt-lite\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\ppt-lite"; Filename: "{app}\{#AppExe}"; Comment: "启动 ppt-lite（浏览器访问 http://127.0.0.1:8765）"
Name: "{autodesktop}\ppt-lite"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon
Name: "{group}\卸载 ppt-lite"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#AppExe}"; Description: "立即启动 ppt-lite"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{localappdata}\ppt-lite\data\tmp"
Type: filesandordirs; Name: "{localappdata}\ppt-lite\data\trash"

[Code]
function InitializeSetup(): Boolean;
var
  mbRes: Integer;
begin
  mbRes := MsgBox('LibreOffice 不在本安装包内。'#13#10 +
    '若需旧 .ppt 转换与页面预览，请另行安装 LibreOffice（可选，不装也可用，仅渲染降级）。'#13#10 +
    '是否继续安装？', mbConfirmation, MB_YESNO);
  Result := (mbRes = IDYES);
end;
