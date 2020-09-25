' Script to Reset last logged in user name for Lync and SfB Client
'
' Version 1.0 - 2018 December 5
'
' Lync 2010:
'
' Remove: HKCU\Software\Microsoft\Shared\Ucclient\ServerSipUri
'
' Lync 2013:
'
' Remove: HKCU\Software\Microsoft\Office\15.0\Lync\ServerSipUri
' Remove: HKCU\Software\Microsoft\Office\15.0\Lync\WindowsAccountSipUri
' Set:    HKCU\Software\Microsoft\Office\15.0\Lync\SavePassword = (REG_DWORD) 0
' Delete: %USERPROFILE%\AppData\Roaming\Microsoft\Office\15.0\Lync\AccountProfiles.dat
'
' Skype for Business 2016:
'
' Remove: HKCU\Software\Microsoft\Office\16.0\Lync\ServerSipUri
' Remove: HKCU\Software\Microsoft\Office\16.0\Lync\WindowsAccountSipUri
' Set:    HKCU\Software\Microsoft\Office\16.0\Lync\SavePassword = (REG_DWORD) 0
' Delete: %USERPROFILE%\AppData\Roaming\Microsoft\Office\16.0\Lync\AccountProfiles.dat

Option Explicit

On Error Resume Next

Const EVENT_SUCCESS      = 0
Const EVENT_ERROR        = 1
Const EVENT_WARNING      = 2
Const EVENT_INFORMATION  = 4

Const REG_HKLM           = &H80000002
Const REG_HKCU           = &H80000001

Const REG_L2010_KEY      = "Software\Microsoft\Shared\Ucclient"
Const REG_L2013_KEY      = "Software\Microsoft\Office\15.0\Lync"
Const REG_S2016_KEY      = "Software\Microsoft\Office\16.0\Lync"

Const DIR_L2013          = "%USERPROFILE%\AppData\Roaming\Microsoft\Office\15.0\Lync"
Const DIR_S2016          = "%USERPROFILE%\AppData\Roaming\Microsoft\Office\16.0\Lync"

'--------------------------------------
' Run-time variables

Dim sEventLog, aLogLevels(4)
Dim oShell, oReg, oFSO, sServerSipUri, sAccountProfiles
Dim sResult

sEventLog  = "Script: " & WScript.ScriptName & vbCrLf & vbCrLf
sResult    = ""

' Event Log Levels
aLogLevels(EVENT_SUCCESS)     = "Success"
aLogLevels(EVENT_ERROR)       = "  Error"
aLogLevels(EVENT_WARNING)     = "Warning"
aLogLevels(EVENT_INFORMATION) = "   Info"

' Create system objects
Set oShell = CreateObject ("WScript.Shell")                     : If oShell Is Nothing Then WScript.Quit (101)
Set oReg   = GetObject ("winmgmts://./root/default:StdRegProv") : If oReg   Is Nothing Then LogEventExit EVENT_ERROR, "Could not access Registry", 102
Set oFSO   = CreateObject ("Scripting.FileSystemObject")        : If oFSO   Is Nothing Then LogEventExit EVENT_ERROR, "Could not open File System", 103

' Check for Lync 2010, and act if it was there
oReg.GetStringValue REG_HKCU, REG_L2010_KEY, "ServerSipUri", sServerSipUri
If 8 = VarType (sServerSipUri) Then

 oReg.DeleteValue REG_HKCU, REG_L2010_KEY, "ServerSipUri"
 sResult = sResult & vbCrLf & "Deleted ""HKCU\" & REG_L2010_KEY & "\ServerSipUri"""

End If

' Check for Lync 2013, and act if it was there
oReg.GetStringValue REG_HKCU, REG_L2013_KEY, "ServerSipUri", sServerSipUri
If 8 = VarType (sServerSipUri) Then

 oReg.DeleteValue REG_HKCU, REG_L2013_KEY, "ServerSipUri"
 sResult = sResult & vbCrLf & "Deleted ""HKCU\" & REG_L2013_KEY & "\ServerSipUri"""

 oReg.DeleteValue REG_HKCU, REG_L2013_KEY, "WindowsAccountSipUri"
 sResult = sResult & vbCrLf & "Deleted ""HKCU\" & REG_L2013_KEY & "\WindowsAccountSipUri"""

 oReg.SetDWORDValue REG_HKCU, REG_L2013_KEY, "SavePassword", 0
 sResult = sResult & vbCrLf & "Set ""HKCU\" & REG_L2013_KEY & "\SavePassword"" into 0 (DWORD)"

 sAccountProfiles = oShell.ExpandEnvironmentStrings (DIR_L2013 & "\AccountProfiles.dat")
 If oFSO.FileExists (sAccountProfiles) Then
  oFSO.DeleteFile sAccountProfiles, True
  sResult = sResult & vbCrLf & "Deleted """ & sAccountProfiles & """"
 End If

End If

' Check for Skype for Business 2016, and act if it was there
oReg.GetStringValue REG_HKCU, REG_S2016_KEY, "ServerSipUri", sServerSipUri
If 8 = VarType (sServerSipUri) Then

 oReg.DeleteValue REG_HKCU, REG_S2016_KEY, "ServerSipUri"
 sResult = sResult & vbCrLf & "Deleted ""HKCU\" & REG_S2016_KEY & "\ServerSipUri"""

 oReg.DeleteValue REG_HKCU, REG_S2016_KEY, "WindowsAccountSipUri"
 sResult = sResult & vbCrLf & "Deleted ""HKCU\" & REG_S2016_KEY & "\WindowsAccountSipUri"""

 oReg.SetDWORDValue REG_HKCU, REG_S2016_KEY, "SavePassword", 0
 sResult = sResult & vbCrLf & "Set ""HKCU\" & REG_S2016_KEY & "\SavePassword"" into 0 (DWORD)"

 sAccountProfiles = oShell.ExpandEnvironmentStrings (DIR_S2016 & "\AccountProfiles.dat")
 If oFSO.FileExists (sAccountProfiles) Then
  On Error Resume Next
  oFSO.DeleteFile sAccountProfiles, True
  On Error GoTo 0
  sResult = sResult & vbCrLf & "Deleted """ & sAccountProfiles & """"
 End If

End If

If Len (sResult) > 2 Then LogEventExit EVENT_SUCCESS, Mid (sResult, 3), 1

Set oFSO   = Nothing
Set oReg   = Nothing
Set oShell = Nothing

WScript.Quit (0)

'--------------------------------------
Sub LogEventExit (iLevel, sMessage, iExit)

 oShell.LogEvent iLevel, sEventLog & sMessage

 If iExit > 0 Then WScript.Quit (iExit)

End Sub
