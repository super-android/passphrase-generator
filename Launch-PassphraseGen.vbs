Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")

' Get the folder this .vbs file lives in
scriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
ps1Path   = scriptDir & "\PassphraseGen.ps1"

' Launch PowerShell hidden - no console window, no blue flash
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & ps1Path & """", 0, False
