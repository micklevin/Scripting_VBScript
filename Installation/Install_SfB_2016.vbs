' Script to install the application
'
' Version 1.1 - 2016 September 23
'  Supports both MSI and EXE installers
' Version 1.2 - 2017 May 16
'  Supports version level
' Version 1.2.1 - 2016 June 8
'  Supports error code decryption

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
Const sMsiExec           = "%WINDIR%\System32\MSIEXEC.EXE"

'--------------------------------------
' Configuration variables
Dim aProductRegs, iProductRegs, sProductRegKey, sProductRegVal, sProductRegData

' Configuration constants
'

' If installing Application as Microsoft Office component:

Const sApplicationName   = "Skype for Business 2016"


' Registry keys which are present if an Office application is already installed
aProductRegs             = Array ("SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Lync\InstallationDirectory", _
                                  "SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Lync\InstallationDirectory", _
                                  "SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Lync\Capabilities\ApplicationName", _
                                  "SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Lync\Capabilities\ApplicationName")

' Else if installing standalone software:
'
' Name of the product, as it appears in Windows Uninstaller.
' This name will be looked for to confirm the need to run the Installer
' This is the value if "DisplayName" of the Registry Key, located under sKeyUninstall

Const sProductName       = ""
Const sProductVersion    = ""
' A registry path to either 64 or 32 bit branch to search for software installation data
Const sKeyUninstall      = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"	' x64
'                          "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"		' x86

' UNC path to the Installer. Make sure it's well-distributed among site where clients who need it reside.
Const sInstallPath       = "\\domain.local\DFS\IT\Applications\Microsoft Office\Skype for Business 2016\Setup x86"

' MSI version

Const sInstallMsi_x64    = ""
Const sInstallMsi_x86    = ""
Const sTransformMsi_x64  = ""
Const sTransformMsi_x86  = ""

' Exe version

Const sInstallExe_x64    = "setup.exe"
Const sInstallExe_x86    = "setup.exe"

' Additional parameters

Const sInstallParams_x64 = ""
Const sInstallParams_x86 = ""

'--------------------------------------
' Run-time variables

Dim oReg, oShell, oFSO, oExec
Dim sEventLog, iExit
Dim bInstalled, bTimeOut, dLimit
Dim sDisplayName, sDisplayVersion
Dim sInstallExe, sInstallParams, sInstallCommand
Dim aSubKeys, sSubKey, aValues, aTypes
Dim i, j

Set oFSO   = CreateObject ("Scripting.FileSystemObject")
If oFSO Is Nothing Then WScript.Quit (101)

Set oShell = CreateObject ("WScript.Shell")
If oShell Is Nothing Then WScript.Quit (102)

Set oReg = GetObject ("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
If oReg Is Nothing Then
 LogEvent EVENT_ERROR, "Could not open Registry object"
 WScript.Quit (103)
End If

' Determine the platform

If 64 = GetObject ("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth Then
 If "" = sInstallMsi_x64 Then
  sInstallExe    = sInstallPath & "\" & sInstallExe_x64
  sInstallParams = sInstallParams_x64
 Else
  sInstallExe    = sMsiExec
  sInstallParams = "/i """ & sInstallPath & "\" & sInstallMsi_x64 & """"
  If "" <> sTransformMsi_x64  Then sInstallParams = sInstallParams & " TRANSFORMS=""" & sInstallPath & "\" & sTransformMsi_x64 & """"
  If "" <> sInstallParams_x64 Then sInstallParams = sInstallParams & " " & sInstallParams_x64
 End If
Else
 If "" = sInstallMsi_x86 Then
  sInstallExe    = sInstallPath & "\" & sInstallExe_x86
  sInstallParams = sInstallParams_x86
 Else
  sInstallExe    = sMsiExec
  sInstallParams = "/i """ & sInstallPath & "\" & sInstallMsi_x86 & """"
  If "" <> sTransformMsi_x86  Then sInstallParams = sInstallParams & " TRANSFORMS=""" & sInstallPath & "\" & sTransformMsi_x86 & """"
  If "" <> sInstallParams_x86 Then sInstallParams = sInstallParams & " " & sInstallParams_x86
 End If
End If

sInstallExe    = Trim (oShell.ExpandEnvironmentStrings (sInstallExe))
sInstallParams = Trim (oShell.ExpandEnvironmentStrings (sInstallParams))

If "" = sInstallParams Then sInstallCommand = sInstallExe Else sInstallCommand = sInstallExe & " " & Trim (sInstallParams)

' Limit of time to wait for the installation to finished (15 minutes)

dLimit     = DateAdd ("n", 25, Now)
bInstalled = False
sEventLog  = "Script: " & WScript.ScriptName & vbCrLf & vbCrLf

'--------------------------------------
' Process

'--------------------------------------
' Check for application's presence via specific Registry Keys
If sProductName = "" Then
 If IsArray (aProductRegs) Then
  iProductRegs = UBound (aProductRegs)
  For i = 0 to iProductRegs
   j = InStrRev (aProductRegs(i), "\")
   If j > 2 And j < Len (aProductRegs(i)) Then
    sProductRegKey = Left (aProductRegs(i), j - 1)
    sProductRegVal = Mid (aProductRegs(i), j + 1)
    oReg.GetStringValue HKEY_LOCAL_MACHINE, sProductRegKey, sProductRegVal, sProductRegData
    If 8 = VarType (sProductRegData) Then
     LogEvent EVENT_INFORMATION, "The Registry contains record under value ""[HKLM\" & sProductRegKey & "]"", """ & sProductRegVal & """"
     bInstalled = True
     Exit For
    End If
   Else
    LogEvent EVENT_WARNING, "The Registry configuraion is wrong with line # " & (i + 1) & ": """ & aProductRegs(i) & """, (" & j & ")"
    WScript.Quit (106)
   End If
  Next
 Else
  LogEvent EVENT_ERROR, "The Registry configuraion is not defined"
  WScript.Quit (105)
 End If
Else
' Check if an application with that name is already installed
 oReg.EnumKey HKEY_LOCAL_MACHINE, sKeyUninstall, aSubKeys
 If Not IsArray (aSubKeys) Then
  LogEvent EVENT_ERROR, "Registry does not contain keys under """ & sKeyUninstall & """"
  WScript.Quit (104)
 End If

 If DEBUG_IT Then WScript.Echo "> Loaded keys under """ & sKeyUninstall & """"

 For Each sSubKey In aSubKeys
  oReg.EnumValues HKEY_LOCAL_MACHINE, sKeyUninstall & "\" & sSubKey, aValues, aTypes
  If (8191 < VarType (aValues)) And (8191 < VarType (aTypes)) Then
   i = UBound (aValues)
   For j = 0 To i
    If "DISPLAYNAME" = UCase (aValues(j)) And REG_SZ = aTypes(j) Then
     oReg.GetStringValue HKEY_LOCAL_MACHINE, sKeyUninstall & "\" & sSubKey, "DisplayName", sDisplayName
     If 8 = VarType (sDisplayName) Then
      If sProductName = sDisplayName Then
       If DEBUG_IT Then WScript.Echo "> Product """ & sProductName & """ found in Registry"
       oReg.GetStringValue HKEY_LOCAL_MACHINE, sKeyUninstall & "\" & sSubKey, "DisplayVersion", sDisplayVersion
       If 8 = VarType (sDisplayVersion) Then
        If DEBUG_IT Then WScript.Echo "> Version """ & sDisplayVersion & """ found in Registry"
        If sProductVersion = sDisplayVersion Then
         bInstalled = True
        Else
         If 1 = StrComp (sDisplayVersion, sProductVersion, vbTextCompare) Then
          bInstalled = True
         End If
        End If
       End If
       Exit For
      End If
     End If
    End If
   Next
   If bInstalled Then Exit For
  End If
 Next
End If

Set oReg = Nothing

'--------------------------------------
' If the application was installed before, log that and exit with code 100
' If the application is not installed yet, do it now

If bInstalled Then
 If sProductName = "" Then
  LogEvent EVENT_INFORMATION, "The product """ & sApplicationName & """ already installed"
 Else
  LogEvent EVENT_INFORMATION, "The product """ & sProductName & """ (current version: " & sDisplayVersion & ", required: " & sProductVersion & ") already installed"
 End If
 iExit = 100
Else
 If oFSO.FileExists (sInstallExe) Then

  If DEBUG_IT Then
   If sProductName = "" Then
    WScript.Echo "The application """ & sApplicationName & """ seems not installed"
    WScript.Echo "The command line to install """ & sApplicationName & """ is:" & vbCrLf & vbCrLf & _
                 sInstallCommand & vbCrLf
   Else
    WScript.Echo "The product """ & sProductName & """ (current version: " & sDisplayVersion & ", required: " & sProductVersion & ") seems not installed."
    WScript.Echo "The command line to install """ & sProductName & """ is:" & vbCrLf & vbCrLf & _
                 sInstallCommand & vbCrLf
   End If
   iExit = 0
  Else
   If sProductName = "" Then
    LogEvent EVENT_INFORMATION, "Starting the installation of """ & sApplicationName & """ from """ & sInstallCommand & """"
   Else
    If "" <> sDisplayVersion Then
     LogEvent EVENT_INFORMATION, "Starting the installation of """ & sProductName & """ as """ & sInstallCommand & """ (previous version: " & sDisplayVersion & ", required: " & sProductVersion & ")"
    Else
     LogEvent EVENT_INFORMATION, "Starting the installation of """ & sProductName & """ as """ & sInstallCommand & """"
    End If
   End If

   ' https://msdn.microsoft.com/en-us/library/d5fk67ky
   Set oExec = oShell.Exec (sInstallCommand)

   bTimeOut = False
   Do While oExec.Status = 0
    If Now < dLimit Then
     WScript.Sleep 500
    Else
     On Error Resume Next
     oExec.Terminate
     bTimeOut = True
    End If
   Loop

   If bTimeOut Then
    If sProductName = "" Then
     LogEvent EVENT_ERROR, "The installation of """ & sApplicationName & """ did not finish in time."
    Else
     LogEvent EVENT_ERROR, "The installation of """ & sProductName & """ did not finish in time."
    End If
   Else
    If sProductName = "" Then
     LogEvent EVENT_INFORMATION, "The installation of """ & sApplicationName & """ finished with code: " & oExec.ExitCode
    Else
     LogEvent EVENT_INFORMATION, "The installation of """ & sProductName & """ finished with code: " & iExit & vbCrLf & vbCrLf & "Code description:" & vbCrLf & GetMsiexecReturnCode (oExec.ExitCode)
    End If
    iExit = oExec.ExitCode
   End If
  End If
 Else
  LogEvent EVENT_ERROR, "The installation file """ & sInstallExe & """ does not exist."
  iExit = 0
 End If
End If

'--------------------------------------

Set oShell = Nothing
Set oFSO = Nothing

WScript.Quit (iExit)

'--------------------------------------
Sub LogEvent (ByVal iSeverity, ByRef sMessage)

 If DEBUG_IT Then
  WScript.Echo "> " & sMessage
 Else
  oShell.LogEvent iSeverity, sEventLog & sMessage
 End If

End Sub

'--------------------------------------
Function GetMsiexecReturnCode (ByVal iCode)

 GetMsiexecReturnCode = "N/A"

 Select Case iCode
 Case    0 GetMsiexecReturnCode = "The action completed successfully"
 Case   13 GetMsiexecReturnCode = "The data is invalid"
 Case   87 GetMsiexecReturnCode = "One of the parameters was invalid"
 Case  120 GetMsiexecReturnCode = "This value is returned when a custom action attempts to call a function that cannot be called from custom actions. The function returns the value ERROR_CALL_NOT_IMPLEMENTED. Available beginning with Windows Installer version 3.0"
 Case 1259 GetMsiexecReturnCode = "If Windows Installer determines a product may be incompatible with the current operating system, it displays a dialog box informing the user and asking whether to try to install anyway. This error code is returned if the user chooses not to try the installation"
 Case 1601 GetMsiexecReturnCode = "The Windows Installer service could not be accessed. Contact your support personnel to verify that the Windows Installer service is properly registered"
 Case 1602 GetMsiexecReturnCode = "The user cancels installation"
 Case 1603 GetMsiexecReturnCode = "A fatal error occurred during installation"
 Case 1604 GetMsiexecReturnCode = "Installation suspended, incomplete"
 Case 1605 GetMsiexecReturnCode = "This action is only valid for products that are currently installed"
 Case 1606 GetMsiexecReturnCode = "The feature identifier is not registered"
 Case 1607 GetMsiexecReturnCode = "The component identifier is not registered"
 Case 1608 GetMsiexecReturnCode = "This is an unknown property"
 Case 1609 GetMsiexecReturnCode = "The handle is in an invalid state"
 Case 1610 GetMsiexecReturnCode = "The configuration data for this product is corrupt. Contact your support personnel"
 Case 1611 GetMsiexecReturnCode = "The component qualifier not present"
 Case 1612 GetMsiexecReturnCode = "The installation source for this product is not available. Verify that the source exists and that you can access it"
 Case 1613 GetMsiexecReturnCode = "This installation package cannot be installed by the Windows Installer service. You must install a Windows service pack that contains a newer version of the Windows Installer service"
 Case 1614 GetMsiexecReturnCode = "The product is uninstalled"
 Case 1615 GetMsiexecReturnCode = "The SQL query syntax is invalid or unsupported"
 Case 1616 GetMsiexecReturnCode = "The record field does not exist"
 Case 1618 GetMsiexecReturnCode = "Another installation is already in progress. Complete that installation before proceeding with this install"
 Case 1619 GetMsiexecReturnCode = "This installation package could not be opened. Verify that the package exists and is accessible, or contact the application vendor to verify that this is a valid Windows Installer package"
 Case 1620 GetMsiexecReturnCode = "This installation package could not be opened. Contact the application vendor to verify that this is a valid Windows Installer package"
 Case 1621 GetMsiexecReturnCode = "There was an error starting the Windows Installer service user interface. Contact your support personnel"
 Case 1622 GetMsiexecReturnCode = "There was an error opening installation log file. Verify that the specified log file location exists and is writable"
 Case 1623 GetMsiexecReturnCode = "This language of this installation package is not supported by your system"
 Case 1624 GetMsiexecReturnCode = "There was an error applying transforms. Verify that the specified transform paths are valid"
 Case 1625 GetMsiexecReturnCode = "This installation is forbidden by system policy. Contact your system administrator"
 Case 1626 GetMsiexecReturnCode = "The function could not be executed"
 Case 1627 GetMsiexecReturnCode = "The function failed during execution"
 Case 1628 GetMsiexecReturnCode = "An invalid or unknown table was specified"
 Case 1629 GetMsiexecReturnCode = "The data supplied is the wrong type"
 Case 1630 GetMsiexecReturnCode = "Data of this type is not supported"
 Case 1631 GetMsiexecReturnCode = "The Windows Installer service failed to start. Contact your support personnel"
 Case 1632 GetMsiexecReturnCode = "The Temp folder is either full or inaccessible. Verify that the Temp folder exists and that you can write to it"
 Case 1633 GetMsiexecReturnCode = "This installation package is not supported on this platform. Contact your application vendor"
 Case 1634 GetMsiexecReturnCode = "Component is not used on this machine"
 Case 1635 GetMsiexecReturnCode = "This patch package could not be opened. Verify that the patch package exists and is accessible, or contact the application vendor to verify that this is a valid Windows Installer patch package"
 Case 1636 GetMsiexecReturnCode = "This patch package could not be opened. Contact the application vendor to verify that this is a valid Windows Installer patch package"
 Case 1637 GetMsiexecReturnCode = "This patch package cannot be processed by the Windows Installer service. You must install a Windows service pack that contains a newer version of the Windows Installer service"
 Case 1638 GetMsiexecReturnCode = "Another version of this product is already installed. Installation of this version cannot continue. To configure or remove the existing version of this product, use Add/Remove Programs in Control Panel"
 Case 1639 GetMsiexecReturnCode = "Invalid command line argument. Consult the Windows Installer SDK for detailed command-line help"
 Case 1640 GetMsiexecReturnCode = "The current user is not permitted to perform installations from a client session of a server running the Terminal Server role service"
 Case 1641 GetMsiexecReturnCode = "The installer has initiated a restart. This message is indicative of a success"
 Case 1642 GetMsiexecReturnCode = "The installer cannot install the upgrade patch because the program being upgraded may be missing or the upgrade patch updates a different version of the program. Verify that the program to be upgraded exists on your computer and that you have the correct upgrade patch"
 Case 1643 GetMsiexecReturnCode = "The patch package is not permitted by system policy"
 Case 1644 GetMsiexecReturnCode = "One or more customizations are not permitted by system policy"
 Case 1645 GetMsiexecReturnCode = "Windows Installer does not permit installation from a Remote Desktop Connection"
 Case 1646 GetMsiexecReturnCode = "The patch package is not a removable patch package. Available beginning with Windows Installer version 3.0"
 Case 1647 GetMsiexecReturnCode = "The patch is not applied to this product. Available beginning with Windows Installer version 3.0"
 Case 1648 GetMsiexecReturnCode = "No valid sequence could be found for the set of patches. Available beginning with Windows Installer version 3.0"
 Case 1649 GetMsiexecReturnCode = "Patch removal was disallowed by policy. Available beginning with Windows Installer version 3.0"
 Case 1650 GetMsiexecReturnCode = "The XML patch data is invalid. Available beginning with Windows Installer version 3.0"
 Case 1651 GetMsiexecReturnCode = "Administrative user failed to apply patch for a per-user managed or a per-machine application that is in advertise state. Available beginning with Windows Installer version 3.0"
 Case 1652 GetMsiexecReturnCode = "Windows Installer is not accessible when the computer is in Safe Mode. Exit Safe Mode and try again or try using System Restore to return your computer to a previous state. Available beginning with Windows Installer version 4.0"
 Case 1653 GetMsiexecReturnCode = "Could not perform a multiple-package transaction because rollback has been disabled. Multiple-Package Installations cannot run if rollback is disabled. Available beginning with Windows Installer version 4.5"
 Case 1654 GetMsiexecReturnCode = "The app that you are trying to run is not supported on this version of Windows. A Windows Installer package, patch, or transform that has not been signed by Microsoft cannot be installed on an ARM computer"
 Case 3010 GetMsiexecReturnCode = "A restart is required to complete the install. This message is indicative of a success. This does not include installs where the ForceReboot action is run"
 End Select

End Function
