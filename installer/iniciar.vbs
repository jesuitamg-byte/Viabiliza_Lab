' Abre o PNBOX Academico (janela do Python minimizada)
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
d = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = d
sh.Run """" & d & "\python\python.exe"" """ & d & "\app_pnbox.py""", 7, False
