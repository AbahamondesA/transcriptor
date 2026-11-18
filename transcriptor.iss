; Instalador del Transcriptor (Inno Setup 6)
[Setup]
AppId={{430A47D1-2ED8-4926-A80E-F9F099CC4868}
AppName=Transcriptor
AppVersion=1.0
AppPublisher=Alejandro Bahamondes
AppComments=Transcribe archivos, enlaces y reuniones en vivo; traduce inglés-español; crea videos subtitulados.
DefaultDirName={localappdata}\Programs\Transcriptor
DefaultGroupName=Transcriptor
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=instalador
OutputBaseFilename=Transcriptor-Setup
SetupIconFile=transcriptor.ico
UninstallDisplayIcon={app}\Transcriptor.exe
UninstallDisplayName=Transcriptor
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes

[Languages]
Name: "es"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "escritorio"; Description: "Crear acceso directo en el escritorio"
Name: "enviara"; Description: "Agregar Transcriptor al menú del clic derecho (Enviar a)"

[Files]
Source: "dist\Transcriptor\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "LEEME.txt"; DestDir: "{app}"; Flags: ignoreversion isreadme

[Icons]
Name: "{autoprograms}\Transcriptor"; Filename: "{app}\Transcriptor.exe"
Name: "{autodesktop}\Transcriptor"; Filename: "{app}\Transcriptor.exe"; Tasks: escritorio
Name: "{usersendto}\Transcriptor"; Filename: "{app}\Transcriptor.exe"; Tasks: enviara
Name: "{autoprograms}\Desinstalar Transcriptor"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\Transcriptor.exe"; Description: "Abrir Transcriptor"; Flags: nowait postinstall skipifsilent
