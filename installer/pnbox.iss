; Instalador do ViabilizaLab (gerado automaticamente pelo GitHub Actions)
#ifndef ProgVersao
  #define ProgVersao "0"
#endif
#ifndef Versao
  #define Versao "modelo"
#endif
#if Versao == "modelo"
  #define Nome "ViabilizaLab - Modelo FitBox"
  #define Arq "ViabilizaLab-Setup-Modelo-FitBox"
  #define Pasta "ViabilizaLab Modelo"
  #define Id "PNBOXAcademicoModelo"
  #define NomeAntigo "PNBOX Acadêmico - Modelo FitBox"
#else
  #define Nome "ViabilizaLab - Em branco"
  #define Arq "ViabilizaLab-Setup-Em-Branco"
  #define Pasta "ViabilizaLab"
  #define Id "PNBOXAcademicoEmBranco"
  #define NomeAntigo "PNBOX Acadêmico - Em branco"
#endif

[Setup]
AppId={#Id}
AppName={#Nome}
AppVersion={#ProgVersao}
AppVerName={#Nome} {#ProgVersao}
AppPublisher=ViabilizaLab - UFVJM
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

[InstallDelete]
; atalhos da versão anterior, quando o programa se chamava PNBOX Acadêmico
Type: files; Name: "{autodesktop}\{#NomeAntigo}.lnk"
Type: files; Name: "{autoprograms}\{#NomeAntigo}.lnk"

[Icons]
Name: "{autodesktop}\{#Nome}"; Filename: "{sys}\wscript.exe"; Parameters: """{app}\iniciar.vbs"""; WorkingDir: "{app}"; IconFilename: "{app}\pnbox.ico"
Name: "{autoprograms}\{#Nome}"; Filename: "{sys}\wscript.exe"; Parameters: """{app}\iniciar.vbs"""; WorkingDir: "{app}"; IconFilename: "{app}\pnbox.ico"

[Run]
Filename: "{sys}\wscript.exe"; Parameters: """{app}\iniciar.vbs"""; Description: "Abrir o ViabilizaLab agora"; Flags: postinstall nowait skipifsilent
