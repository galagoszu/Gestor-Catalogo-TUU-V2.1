
#define MyAppName "Gestor de Catálogo TUU"
#define MyAppVersion "2.1"
#define MyAppExeName "GestorTUU.exe"

[Setup]
AppId={{B78E21F4-4F3B-4A7E-9D85-2D7E1A6C092B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=Gestor TUU
DefaultDirName={localappdata}\Programs\GestorTUU
DefaultGroupName=Gestor de Catálogo TUU V2.1
OutputDir=instaladores
OutputBaseFilename=GestorTUU_Setup_v2.1_CORREGIDO
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesAllowed=x64compatible

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear un acceso directo en el escritorio"; GroupDescription: "Accesos directos:"; Flags: checkedonce

[Files]
Source: "dist\GestorTUU_v2.1\GestorTUU\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Gestor de Catálogo TUU V2.1"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Gestor de Catálogo TUU V2.1"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir Gestor de Catálogo TUU V2.1"; Flags: nowait postinstall skipifsilent
