; Inno Setup 脚本
; 生成 Windows 安装程序 .exe

#define MyAppName      "批签发证明拆分工具"
#define MyAppPublisher "批签发工具"
#define MyAppURL       "https://github.com/YOUR_ORG/batch_cert_splitter"
#define MyAppExeName   "批签发拆分工具.exe"
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif

[Setup]
AppId={{A3B2C1D0-E4F5-6789-ABCD-EF0123456789}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=..\..\dist
OutputBaseFilename=批签发拆分工具_v{#MyAppVersion}_Windows_x64
Compression=lzma2/ultra64
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
WizardStyle=modern
DisableDirPage=no
PrivilegesRequired=lowest
; 允许非管理员安装到用户目录

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务："; Flags: unchecked

[Files]
Source: "..\..\dist\批签发拆分工具\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}";       Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}";  Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent
