Option Explicit

' Launches the media-pipeline watcher with no visible window. Resolves its own
' folder, so the whole app folder can be copied/moved anywhere.

Dim fso, shell, scriptDir, scriptPath, pwshPath, command
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = fso.BuildPath(scriptDir, "watch-media.ps1")

' Prefer PowerShell 7 (enables parallel image processing); fall back to the
' portable build, then to Windows PowerShell 5.1.
pwshPath = "C:\Program Files\PowerShell\7\pwsh.exe"
If Not fso.FileExists(pwshPath) Then pwshPath = "C:\Tools\pwsh\pwsh.exe"
If Not fso.FileExists(pwshPath) Then pwshPath = "powershell.exe"

command = """" & pwshPath & """ -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """"

' 0 = hidden window, False = do not wait.
shell.Run command, 0, False
