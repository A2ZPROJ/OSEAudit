[Setup]
AppName=OSEAudit
AppVersion=1.0
AppPublisher=A2Z Projetos
DefaultDirName={autopf}\A2Z Projetos\OSEAudit
DefaultGroupName=A2Z Projetos\OSEAudit
AllowNoIcons=yes
OutputDir=.
OutputBaseFilename=OSEAudit_Setup
SetupIconFile=assets\oseaudit.ico
UninstallDisplayIcon={app}\OSEAudit.exe
Compression=lzma2/fast
SolidCompression=no
WizardStyle=modern
PrivilegesRequiredOverridesAllowed=dialog
MinVersion=10.0
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Area de Trabalho"; GroupDescription: "Icones adicionais:"

[Files]
Source: "dist\OSEAudit\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\OSEAudit"; Filename: "{app}\OSEAudit.exe"
Name: "{group}\Desinstalar OSEAudit"; Filename: "{uninstallexe}"
Name: "{autodesktop}\OSEAudit"; Filename: "{app}\OSEAudit.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\OSEAudit.exe"; Description: "Abrir OSEAudit agora"; Flags: nowait postinstall skipifsilent
