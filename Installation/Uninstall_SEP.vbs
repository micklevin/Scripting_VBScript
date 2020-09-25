' Script to un-install the application
'
' Version 1.2 - 2016 October 17
'  Supports Alternative Uninstall (by product UID, rather by what is in Registry)
'  Supports post-installation command
'  Exit codes:
'   < 99 = exit code from MSIEXEC.EXE
'     99 = Debug mode, runs nothing, just reports steps done
'    100 = Product is not installed
'    103 = Registry does not contain keys under REG_UNINSTALL
'    104 = Registry does not contain data UninstallString
'    105 = Empty value of data UninstallString
'    106 = Expanding post-uninstall EXE name returns empty string
'    107 = Post-install EXE file does not exist

Option Explicit

Const DEBUG_IT = True

If Not DEBUG_IT Then On Error Resume Next

'--------------------------------------
' System constants
'

Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_SZ             = 1
Const REG_EXPAND_SZ      = 2
Const REG_BINARY         = 3
Const REG_DWORD          = 4
Const REG_MULTI_SZ       = 7
Const EVENT_SUCCESS      = 0
Const EVENT_ERROR        = 1
Const EVENT_WARNING      = 2
Const EVENT_INFORMATION  = 4

'--------------------------------------
' Configuration constants

' Name of the product, as it appears in Windows Uninstaller.
' This name will be looked for to confirm the need to run the Installer
' This is the value if "DisplayName" of the Registry Key, located under REG_UNINSTALL

Const PRODUCT_NAME   = "Symantec Endpoint Protection"

' MAke sure you know which branch of Registry (x86 or x64) to look for the uninstallation data

Const REG_UNINSTALL  = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" ' x86 in this case

' If the uninstallation must be done by UID, set this to True
' Otherwise it will find the Uninstallation String in registry, and will run that string instead. Which should be a default behaviour.

Const CUSTOM_INSTALL = True

' In this case, we want another antivirus - Cylance - re-register itself, in case it was already installed, co-existing with Symantec.
' Which is logical, because Cylance can co-exist with another anti-virus, and we want to make sure that after Symantec is gone, we still have a/v protection immediately.
' Which is why Cylance should have been deployed ahead of time.
' Re-registering Cylance would make it recognized by Windows as official 3rd party antivirus.
' While when it was co-existing with SEP, it was ... just ... there.
' And after Symantec is gone, there is a chance Windows would not see Cylance as a first class citizen. Hence we better re-register.

Const POST_EXE       = "%programfiles%\Cylance\Desktop\CylanceSvc.exe"	' Make it empty to skip running post-uninstall executable; No double quotes
Const POST_EXE_PARAM = " /register /enable"				' Keep the leading space, if any

'--------------------------------------
' Run-time variables

Dim oReg, oShell
Dim sEventLog, iExit
Dim sUninstallCommand
Dim aSubKeys, sSubKey, aValues, aTypes
Dim i, j

sEventLog  = "Script: " & WScript.ScriptName & vbCrLf & vbCrLf

Set oShell = CreateObject ("WScript.Shell")
If oShell Is Nothing Then WScript.Quit (101)

Set oReg = GetObject ("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
If oReg Is Nothing Then
 LogEvent EVENT_ERROR, "Could not open Registry object"
 Set oShell = Nothing
 WScript.Quit (102)
End If

' Find the Uninstall Command Line String
sUninstallCommand = RegFindUninstallString ()

If IsNumeric (sUninstallCommand) Then
 iExit = CInt (sUninstallCommand)
Else
 ' Run the Uninstall Command
 LogEvent EVENT_INFORMATION, "Running the following command line to uninstall """ & PRODUCT_NAME & """:" & vbCrLf & vbCrLf & sUninstallCommand
 If DEBUG_IT Then iExit = "99" Else iExit = oShell.Run (sUninstallCommand, 1, True)
 LogEvent EVENT_INFORMATION, "Uninstall command for """ & PRODUCT_NAME & """ finished with code: " & iExit

 ' Run any Post-Uninstall Command
 ShellRunCommand (iExit)
End If

Set oReg   = Nothing
Set oShell = Nothing

WScript.Quit (iExit)

'--------------------------------------
Function RegFindUninstallString ()
 Dim aSubKeys, sSubKey, aValues, aTypes, sDisplayName, sUninstallCommand
 Dim i, j

 ' Default 100, means "not found"
 RegFindUninstallString = 100

 Do
  oReg.EnumKey HKEY_LOCAL_MACHINE, REG_UNINSTALL, aSubKeys
  If Not IsArray (aSubKeys) Then
   LogEvent EVENT_ERROR, "Registry does not contain keys under """ & REG_UNINSTALL & """"
   RegFindUninstallString = 103
   Exit Do
  End If

  For Each sSubKey In aSubKeys
   oReg.EnumValues HKEY_LOCAL_MACHINE, REG_UNINSTALL & "\" & sSubKey, aValues, aTypes

   If (8191 < VarType (aValues)) And (8191 < VarType (aTypes)) Then
    i = UBound (aValues)
    For j = 0 To i
     If "DISPLAYNAME" = UCase (aValues(j)) And REG_SZ = aTypes(j) Then
      oReg.GetStringValue HKEY_LOCAL_MACHINE, REG_UNINSTALL & "\" & sSubKey, "DisplayName", sDisplayName
      If 8 = VarType (sDisplayName) Then
       If PRODUCT_NAME = sDisplayName Then

        If CUSTOM_INSTALL Then
         sUninstallCommand = "MSIEXEC.EXE /X " & sSubKey & " /qn /norestart"
        Else
         oReg.GetStringValue HKEY_LOCAL_MACHINE, REG_UNINSTALL & "\" & sSubKey, "UninstallString", sUninstallCommand
         If 8 <> VarType (sUninstallCommand) Then
          LogEvent EVENT_ERROR, "Registry does not contain value ""UninstallString"" under the key """ & REG_UNINSTALL & "\" & sSubKey & """"
          RegFindUninstallString = 104
          Exit Do
         End If

         If "" = Trim (sUninstallCommand) Then
          LogEvent EVENT_ERROR, "Empty registry value of ""UninstallString"" under the key """ & REG_UNINSTALL & "\" & sSubKey & """"
          RegFindUninstallString = 105
          Exit Do
         End If

         If "MSIEXEC.EXE" = UCase (Left (sUninstallCommand, 11)) Then
          If 0 = InStr (1, UCase (sUninstallCommand), "/QN")        Then sUninstallCommand = sUninstallCommand & " /qn"
          If 0 = InStr (1, UCase (sUninstallCommand), "/NORESTART") Then sUninstallCommand = sUninstallCommand & " /norestart"
         End If
        End If

        RegFindUninstallString = sUninstallCommand
        Exit For
       End If
      End If
     End If
    Next
   End If
  Next
 Loop While False

End Function

'--------------------------------------
Function ShellRunCommand (ByVal iExit)
 Dim oFSO, sCommandLine

 If "" <> POST_EXE Then
  sCommandLine = oShell.ExpandEnvironmentStrings (POST_EXE)

  If "" <> sCommandLine Then
   Set oFSO = CreateObject("Scripting.FileSystemObject")

   If (oFSO.FileExists (sCommandLine)) Then
    sCommandLine = """" & sCommandLine & """" & POST_EXE_PARAM

    LogEvent EVENT_INFORMATION, "Running the following command after uninstalling """ & PRODUCT_NAME & """:" & vbCrLf & vbCrLf & sCommandLine
    If DEBUG_IT Then iExit = "99" Else iExit = oShell.Run (sCommandLine, 1, True)
    LogEvent EVENT_INFORMATION, "The command, which ran after uninstalling """ & PRODUCT_NAME & """, finished with code: " & iExit
   Else
    ShellRunCommand = 107
   End If

   Set oFSO = Nothing
  Else
   ShellRunCommand = 106
  End If
 Else
  ShellRunCommand = iExit
 End If

End Function

'--------------------------------------
Sub LogEvent (iLevel, sMessage)

 If DEBUG_IT Then WScript.Echo iLevel & ": " & sMessage Else oShell.LogEvent iLevel, sEventLog & sMessage

End Sub

'--------------------------------------
Sub LogDebug (sMessage)

 If DEBUG_IT Then WScript.Echo sMessage

End Sub
