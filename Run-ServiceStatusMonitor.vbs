' Run-ServiceStatusMonitor.vbs
' Launches ServiceStatusMonitor.ps1 with no console flash (a bare powershell.exe
' task action flashes a window even with -WindowStyle Hidden when run interactively).
' Argument 1 = mode: Watchdog (default) or Digest.
Option Explicit
Dim mode, fso, shell, scriptDir, psExe, cmd

If WScript.Arguments.Count > 0 Then
    mode = WScript.Arguments(0)
Else
    mode = "Watchdog"
End If

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
psExe = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

Set shell = CreateObject("WScript.Shell")
cmd = """" & psExe & """ -NoProfile -ExecutionPolicy Bypass -File """ & _
      scriptDir & "\ServiceStatusMonitor.ps1"" -Mode " & mode
shell.Run cmd, 0, True
