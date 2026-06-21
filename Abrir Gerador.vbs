' Atalho da área de trabalho aponta para cá.
' Roda o launcher.ps1 totalmente escondido (sem piscar janela preta).
Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(WScript.ScriptFullName)
sh.Run "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & base & "\launcher.ps1""", 0, False
