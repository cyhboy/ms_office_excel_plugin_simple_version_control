Option Explicit

Dim iArgs(2)
Set iArgs(2) = WScript.Arguments
Dim command
command = Unescape(iArgs(0))
command = "cmd /K echo hello"
Dim sec
sec = iArgs(1)
Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")
WScript.Sleep sec
'MsgBox command
'WshShell.Run command, 0, False

Dim WshShellExec
Set WshShellExec = WshShell.Exec(command)

Const WshRunning = 0
Const WshFinished = 1
Const WshFailed = 2

Dim strOutput
Select Case WshShellExec.Status
	Case WshFinished
	strOutput = WshShellExec.StdOut.ReadAll
	Case WshFailed
	strOutput = WshShellExec.StdErr.ReadAll
End Select
'MsgBox strOutput

WshShell.Popup strOutput, 5

Set WshShellExec = Nothing
Set WshShell = Nothing

'MsgBox "Completed " & command & " after " & sec & " seconds"