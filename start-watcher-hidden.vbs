Option Explicit

Dim shell
Dim scriptPath
Dim command

Set shell = CreateObject("WScript.Shell")
scriptPath = "D:\Projects\media-pipeline\watch-media.ps1"
command = "C:\Tools\pwsh\pwsh.exe -NoProfile -ExecutionPolicy Bypass -File " & Chr(34) & scriptPath & Chr(34)

' 0 = hidden window, False = do not wait.
shell.Run command, 0, False
