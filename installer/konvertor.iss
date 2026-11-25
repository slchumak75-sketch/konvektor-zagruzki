#define AppVersion "2.0.0"
[Setup]
AppId={{B9E3E60A-5083-4A4B-A0C0-AC645780AD16}
AppName=Конвертор загрузки
AppVersion={#AppVersion}
DefaultDirName={localappdata}\Programs\Konvertor
DefaultGroupName=Конвертор загрузки
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=Konvertor-Setup
SetupIconFile=..\assets\konvektor.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
UninstallDisplayIcon={app}\Konvertor.exe

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Files]
Source: "..\dist\Konvertor.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Конвертор загрузки"; Filename: "{app}\Konvertor.exe"
Name: "{autodesktop}\Конвертор загрузки"; Filename: "{app}\Konvertor.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; Flags: unchecked

[Run]
Filename: "{app}\Konvertor.exe"; Description: "Запустить Конвертор загрузки"; Flags: nowait postinstall skipifsilent
