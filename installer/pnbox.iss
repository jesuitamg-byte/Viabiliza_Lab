; Instalador do PNBOX Academico (gerado automaticamente pelo GitHub Actions)
#ifndef Versao
  #define Versao "modelo"
#endif
#if Versao == "modelo"
  #define Nome "PNBOX Acadêmico - Modelo FitBox"
  #define Arq "PNBOX-Setup-Modelo-FitBox"
  #define Pasta "PNBOX Academico Modelo"
  #define Id "PNBOXAcademicoModelo"
#else
  #define Nome "PNBOX Acadêmico - Em branco"
  #define Arq "PNBOX-Setup-Em-Branco"
  #define Pasta "PNBOX Academico"
  #define Id "PNBOXAcademicoEmBranco"
#endif

[Setup]
AppId={#Id}
AppName={#Nome}
AppVersion=4.1
AppPublisher=PNBOX Academico - UFVJM
DefaultDirName={localappdata}\{#Pasta}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputBaseFilename={#Arq}
SetupIconFile=pnbox.ico
UninstallDisplayIcon={app}\pnbox.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "pt"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Files]
Source: "..\dist\PNBOX\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{autodesktop}\{#Nome}"; Filename: "{sys}\wscript.exe"; Parameters: """{app}\iniciar.vbs"""; WorkingDir: "{app}"; IconFilename: "{app}\pnbox.ico"
Name: "{autoprograms}\{#Nome}"; Filename: "{sys}\wscript.exe"; Parameters: """{app}\iniciar.vbs"""; WorkingDir: "{app}"; IconFilename: "{app}\pnbox.ico"

[Run]
Filename: "{sys}\wscript.exe"; Parameters: """{app}\iniciar.vbs"""; Description: "Abrir o PNBOX agora"; Flags: postinstall nowait skipifsilent
