#define MyAppName "BMXShell"
#define MyAppVersion "0.1"
#define MyAppPublisher "Fulgen"

[Setup]
AppId={{E76FA325-1BCC-452A-967B-182167E0CD4A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputBaseFilename=bmx_shell_setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
AllowNetworkDrive=no
AllowUNCPath=no
ArchitecturesAllowed=x64compatible and x86compatible
ArchitecturesInstallIn64BitMode=x64compatible
ChangesAssociations=yes
MinVersion=10.0.17763

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{#BinariesDir}\bmx_shell.dll"; DestDir: "{app}"; Flags: ignoreversion regserver 64bit
Source: "{#BinariesDir}\bmx_shell.pdb"; DestDir: "{app}"; Flags: ignoreversion 64bit

[Icons]
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

