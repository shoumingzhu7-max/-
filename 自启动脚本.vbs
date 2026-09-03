' autostart.vbs - watch a list of agent processes, then auto-open the API balance window.
' Auto-detects its own folder: put this file next to API.vbs and it just works.
' Usage: double-click (silent background), or:
'   cscript //nologo "autostart.vbs" selftest
Option Explicit

Dim FSO, shell, wmi
Set FSO = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
Set wmi = GetObject("winmgmts:\\.\root\cimv2")

' ---- Config: agents to watch ----
' Process names WITHOUT ".exe". If ANY of these is running, the window opens.
' "openclaw" is the exe basename for OpenClaw. If its process is NOT openclaw.exe
' (for example it runs as node.exe or a differently-named exe), change it below to
' the actual process name so the watcher can detect it.
Dim TARGET_APPS
TARGET_APPS = Array("codex", "cursor", "claude", "gemini", "continue", "openclaw")

' Own folder - no hard-coded desktop path.
Dim scriptDir
scriptDir = FSO.GetParentFolderName(WScript.ScriptFullName)

' Main program: API.vbs in the same folder.
' "API" + (67E5 8BE2) = filename, kept ASCII-only for cscript compatibility.
Dim balancePath
balancePath = FSO.BuildPath(scriptDir, "API" & ChrW(&H67E5) & ChrW(&H8BE2) & ".vbs")

If WScript.Arguments.Count > 0 Then
  If LCase(WScript.Arguments(0)) = "selftest" Then
    WScript.Echo "script_dir=" & scriptDir
    WScript.Echo "balance_exists=" & CStr(FSO.FileExists(balancePath))
    WScript.Echo "balance=" & balancePath
    Dim ai
    For Each ai In TARGET_APPS
      WScript.Echo "agent[" & ai & "]=" & CStr(CountProcess(ai))
    Next
    WScript.Echo "balance_open=" & CStr(BalanceWindowOpen())
    WScript.Quit 0
  End If
End If

If Not FSO.FileExists(balancePath) Then
  Dim msg
  msg = "Main program not found:" & vbCrLf & balancePath & vbCrLf & vbCrLf
  msg = msg & "Keep this file in the same folder as API.vbs."
  MsgBox msg, vbExclamation, WScript.ScriptName
  WScript.Quit 1
End If

' Seed current state. Starting with running = False means: if any target app is
' already up when the watcher starts, the first loop pass fires the popup once
' (guarded by BalanceWindowOpen), recovering from a startup race where the
' watcher was launched after the target. After that, normal 0->1 edge logic applies.
Dim running, lastTrigger
running = False
lastTrigger = 0

On Error Resume Next
Do While True
  Dim c
  c = CountAnyProcess(TARGET_APPS)
  If Err.Number <> 0 Then
    Err.Clear
    c = 0
  End If

  If c > 0 And Not running Then
    If (Timer - lastTrigger) >= 30 Then
      If Not BalanceWindowOpen() Then
        shell.Run "wscript.exe """ & balancePath & """", 1, False
        lastTrigger = Timer
      End If
    End If
    running = True
  ElseIf c = 0 Then
    running = False
  End If

  WScript.Sleep 3000
Loop

Function CountAnyProcess(apps)
  CountAnyProcess = 0
  Dim ai
  For Each ai In apps
    CountAnyProcess = CountAnyProcess + CountProcess(ai)
  Next
End Function

Function CountProcess(appName)
  CountProcess = 0
  On Error Resume Next
  Dim q
  Set q = wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='" & appName & ".exe'")
  If Err.Number <> 0 Then
    Err.Clear
    Exit Function
  End If
  Dim p
  For Each p In q
    CountProcess = CountProcess + 1
  Next
End Function

Function BalanceWindowOpen()
  BalanceWindowOpen = False
  On Error Resume Next
  Dim q
  Set q = wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='mshta.exe' AND CommandLine LIKE '%APIKeyBalanceSummary.hta%'")
  If Err.Number <> 0 Then
    Err.Clear
    Exit Function
  End If
  Dim p
  For Each p In q
    BalanceWindowOpen = True
    Exit Function
  Next
End Function
