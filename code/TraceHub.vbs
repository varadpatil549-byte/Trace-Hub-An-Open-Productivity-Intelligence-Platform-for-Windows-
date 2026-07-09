Set objShell = CreateObject("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

scriptFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)

objShell.Run _
"powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ _
& scriptFolder & "\TraceHub.ps1""", _
0, False