' Abre o ViabilizaLab (janela do Python oculta; fecha sozinho ao fechar a aba)
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
d = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = d
sh.Run """" & d & "\python\python.exe"" """ & d & "\app_pnbox.py""", 0, False
