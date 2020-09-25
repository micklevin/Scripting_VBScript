' Script to install the application
'
' Version 1.1 - 2016 September 23
'  Supports both MSI and EXE installers
' Version 1.2 - 2017 May 16
'  Supports version level
' Version 1.2.1 - 2016 June 8
'  Supports error code decryption
' Version 1.3 - 2017 June 26
'  Supports registry lookup at both x64 and x86 branches
' Version 2.0 - 2017 June 27
'  Supports chained installation and more steps
' Version 3.0 - 2018 December 18
'  Supports preloading the installation files on local disk
' Version 3.0.1 - 2018 December 21
'  Supports missing the dependency product check

Option Explicit

Const DEBUG_IT = True

If Not DEBUG_IT Then On Error Resume Next

'--------------------------------------
' System constants
'

Const COMMON_APPDATA       = &H23&  ' the second & denotes a long integer

Const HKEY_LOCAL_MACHINE   = &H80000002
Const REG_SZ               = 1
Const REG_EXPAND_SZ        = 2
Const REG_BINARY           = 3
Const REG_DWORD            = 4
Const REG_MULTI_SZ         = 7
Const EVENT_SUCCESS        = 0
Const EVENT_ERROR          = 1
Const EVENT_WARNING        = 2
Const EVENT_INFORMATION    = 4

'--------------------------------------
' Script's constants
'

Const SYS_MSIEXEC          = "%WINDIR%\System32\MSIEXEC.EXE"
Const REG_UNINSTALL_64     = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
Const REG_UNINSTALL_86     = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

Const APP_PRODUCT_NAME     = 0
Const APP_PRODUCT_VERSION  = 1
Const APP_DISTRO_NETWORK   = 2
Const APP_DISTRO_LOCAL     = 3
Const APP_PRODUCT_TYPE     = 4
Const APP_INSTALL_FILE     = 5
Const APP_INSTALL_FILE_64  = 5
Const APP_INSTALL_FILE_86  = 6
Const APP_TRANSFORM        = 7
Const APP_TRANSFORM_64     = 7
Const APP_TRANSFORM_86     = 8
Const APP_INSTALL_SWITCH   = 9
Const APP_PREREQ_ACTION    = 10
Const APP_PREREQ_SRC       = 11
Const APP_PREREQ_DST       = 12

Const APP_DEP_PRODUCT_NAME = 13
Const APP_DEP_PRODUCT_VER  = 14
Const APP_SIZE             = 14

Const ACT_NONE             = 0
Const ACT_COPY             = 1

'--------------------------------------
' Configuration parameters
'

Dim aProduct : Redim aProduct (APP_SIZE)

'--------------------------------------
' Product

' This product's information
aProduct(APP_PRODUCT_NAME)     = "Cisco AnyConnect Secure Mobility Client"
aProduct(APP_PRODUCT_VERSION)  = "4.6.03049"
aProduct(APP_DISTRO_NETWORK)   = "\\domain.local\DFS\IT\Applications\Cisco AnyConnect Umbrella"
aProduct(APP_DISTRO_LOCAL)     = "%APPDATA%\Install Cache"
aProduct(APP_PRODUCT_TYPE)     = "MSI"
aProduct(APP_INSTALL_FILE_64)  = "anyconnect-win-4.6.03049-core-vpn-predeploy-k9.msi"
aProduct(APP_INSTALL_FILE_86)  = "anyconnect-win-4.6.03049-core-vpn-predeploy-k9.msi"
aProduct(APP_TRANSFORM_64)     = ""
aProduct(APP_TRANSFORM_86)     = ""
aProduct(APP_INSTALL_SWITCH)   = "/qn"

' If anyt file needs to be copied over before the installation
aProduct(APP_PREREQ_ACTION)    = ACT_NONE
aProduct(APP_PREREQ_SRC)       = ""
aProduct(APP_PREREQ_DST)       = ""

' If other product is the pre-requisite
aProduct(APP_DEP_PRODUCT_NAME) = ""
aProduct(APP_DEP_PRODUCT_VER)  = ""

'--------------------------------------
' Run-time variables

Dim oReg, oShell, oFSO
Dim oApp, oAppDir, oDirItem, sProgramData
Dim b64
Dim sEventLog, iExit
Dim aLogLevels(4)

' Event Log Levels
aLogLevels(EVENT_SUCCESS)     = "Success"
aLogLevels(EVENT_ERROR)       = "  Error"
aLogLevels(EVENT_WARNING)     = "Warning"
aLogLevels(EVENT_INFORMATION) = "   Info"

' Determine the platform
If 64 = GetObject ("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth Then b64 = True Else b64 = False

' Create system objects
Set oFSO     = CreateObject ("Scripting.FileSystemObject")                                         : If oFSO     Is Nothing Then LogEventExit EVENT_ERROR, "Could not open File System", 101
Set oShell   = CreateObject ("WScript.Shell")                                                      : If oShell   Is Nothing Then LogEventExit EVENT_ERROR, "Could not open Shell", 102
Set oReg     = GetObject ("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv") : If oReg     Is Nothing Then LogEventExit EVENT_ERROR, "Could not open Registry object", 103
Set oApp     = CreateObject ("Shell.Application")                                                  : If oApp     Is Nothing Then LogEventExit EVENT_ERROR, "Could not open Shell Applications", 104
Set oAppDir  = oApp.Namespace (COMMON_APPDATA)                                                     : If oAppDir  Is Nothing Then LogEventExit EVENT_ERROR, "Could not open Shell Applications namespace", 105
Set oDirItem = oAppDir.Self                                                                        : If oDirItem Is Nothing Then LogEventExit EVENT_ERROR, "Could not open Shell Applications directory", 106
sProgramData = oDirItem.Path                                                                       : If sProgramData = ""   Then LogEventExit EVENT_ERROR, "Could not read Applications directory path", 107

'--------------------------------------
' Process the environment

If Not b64 Then
 aProduct(APP_INSTALL_FILE) = aProduct(APP_INSTALL_FILE_86)
 aProduct(APP_TRANSFORM)    = aProduct(APP_TRANSFORM_86)
End If

aProduct(APP_DISTRO_LOCAL) = Trim (Replace (aProduct(APP_DISTRO_LOCAL), "%APPDATA%", sProgramData, 1, -1, vbTextCompare))

sEventLog  = "Script: " & WScript.ScriptName & vbCrLf & vbCrLf

'--------------------------------------
' Exit code
iExit = 0

' Check If Dependency product is not defined, or already installed
If (aProduct(APP_DEP_PRODUCT_NAME) = "") Or _
   (IsProductInstalled (aProduct(APP_DEP_PRODUCT_NAME), aProduct(APP_DEP_PRODUCT_VER))) Then

 ' And If this product is not installed
 If Not IsProductInstalled (aProduct(APP_PRODUCT_NAME), aProduct(APP_PRODUCT_VERSION)) Then

  ' Prepare local files, copy if necessary
  If aProduct(APP_INSTALL_FILE) <> "" Then
   If Not CopyFileIfNewer (aProduct(APP_INSTALL_FILE), aProduct(APP_DISTRO_NETWORK), aProduct(APP_DISTRO_LOCAL)) Then
    LogEventExit EVENT_ERROR, "Could not copy the file """ & aProduct(APP_INSTALL_FILE) & """", 108
   End If
  End If

  If aProduct(APP_TRANSFORM) <> "" Then
   If Not CopyFileIfNewer (aProduct(APP_TRANSFORM), aProduct(APP_DISTRO_NETWORK), aProduct(APP_DISTRO_LOCAL)) Then
    LogEventExit EVENT_ERROR, "Could not copy the file """ & aProduct(APP_TRANSFORM) & """", 109
   End If
  End If

  If aProduct(APP_PREREQ_SRC) <> "" Then
   If Not CopyFileIfNewer (aProduct(APP_PREREQ_SRC), aProduct(APP_DISTRO_NETWORK), aProduct(APP_DISTRO_LOCAL)) Then
    LogEventExit EVENT_ERROR, "Could not copy the file """ & aProduct(APP_PREREQ_SRC) & """", 110
   End If
  End If

  ' Perform installation
  iExit = InstallProduct ()
 Else
  DebugEcho "Product """ & aProduct(APP_PRODUCT_NAME) & """, version """ & aProduct(APP_PRODUCT_VERSION) & """ is already installed."
 End If
Else
 DebugEcho "Product """ & aProduct(APP_DEP_PRODUCT_NAME) & """, version """ & aProduct(APP_DEP_PRODUCT_VER) & """ is not installed."
End If

Set oReg   = Nothing
Set oShell = Nothing
Set oFSO   = Nothing

WScript.Quit (iExit)

'--------------------------------------
Function InstallProduct ()
 Dim oExec
 Dim bTimeOut, dLimit
 Dim sInstallFile, sInstallExe, sInstallParams, sInstallCommand
 Dim iExit

 If "MSI" = aProduct(APP_PRODUCT_TYPE) Then
  sInstallExe    = SYS_MSIEXEC
  sInstallFile   = aProduct(APP_DISTRO_LOCAL) & "\" & aProduct(APP_INSTALL_FILE)
  sInstallParams = "/i """ & sInstallFile & """"
  If "" <> aProduct(APP_TRANSFORM)      Then sInstallParams = sInstallParams & " TRANSFORMS=""" & aProduct(APP_DISTRO_LOCAL) & "\" & aProduct(APP_TRANSFORM) & """"
  If "" <> aProduct(APP_INSTALL_SWITCH) Then sInstallParams = sInstallParams & " " & aProduct(APP_INSTALL_SWITCH)
 Else
  sInstallExe    = aProduct(APP_DISTRO_LOCAL) & "\" & aProduct(APP_INSTALL_FILE)
  sInstallFile   = sInstallExe
  sInstallParams = aProduct(APP_INSTALL_SWITCH)
 End If

 sInstallExe    = Trim (oShell.ExpandEnvironmentStrings (sInstallExe))
 sInstallParams = Trim (oShell.ExpandEnvironmentStrings (sInstallParams))
 sInstallFile   = Trim (oShell.ExpandEnvironmentStrings (sInstallFile))

 ' Check if installation file exists
 If Not oFSO.FileExists (sInstallFile) Then
  LogEvent EVENT_ERROR, "The installation file """ & sInstallFile & """ does not exist."
  InstallProduct = 111
  Exit Function
 End If

 ' Check the pre-requisite
 If ACT_NONE <> aProduct(APP_PREREQ_ACTION) Then
  DebugEcho "Need to execute pre-requisite action: " & aProduct(APP_PREREQ_ACTION)
  iExit = RunPrereq (aProduct(APP_PREREQ_ACTION), aProduct(APP_DISTRO_LOCAL) & "\" & aProduct(APP_PREREQ_SRC), aProduct(APP_PREREQ_DST))
  If iExit > 0 Then
   LogEvent EVENT_ERROR, "The pre-requisite of installation of """ & aProduct(APP_PRODUCT_NAME) & """ returned error: " & iExit
   InstallProduct = 112
   Exit Function
  End If
 End If

 ' Limit of time to wait for the installation to finished (15 minutes)
 dLimit = DateAdd ("n", 15, Now)

 If "" = sInstallParams Then sInstallCommand = sInstallExe Else sInstallCommand = sInstallExe & " " & sInstallParams

 ' Run the installation
 LogEvent EVENT_INFORMATION, "Starting the installation of """ & aProduct(APP_PRODUCT_NAME) & """ as: " & sInstallCommand

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
  LogEvent EVENT_ERROR, "The installation of """ & aProduct(APP_PRODUCT_NAME) & """ did not finish in time."
  InstallProduct = 113
 Else
  InstallProduct = oExec.ExitCode
  LogEvent EVENT_INFORMATION, "The installation of """ & aProduct(APP_PRODUCT_NAME) & """ finished with code: " & InstallProduct & " (" & GetMsiexecReturnCode (InstallProduct) & ")"
 End If

 Set oExec = Nothing

End Function

'--------------------------------------
' Run pre-requisite tasks
' Currently supports only one task: ACT_COPY, which copies file(s)
Function RunPrereq (ByVal iAction, ByRef sSource, ByRef sDestination)
 Dim sDestinationDir

 RunPrereq = 0

 Select Case iAction
 Case ACT_COPY
  sSource          = Trim (oShell.ExpandEnvironmentStrings (sSource))
  sDestination     = Trim (oShell.ExpandEnvironmentStrings (sDestination))
  sDestinationDir  = GetDirFromPath (sDestination)

  ' Check If source file exists
  If Not oFSO.FileExists (sSource) Then
   DebugEcho "Source file does not exist: """ & sSource & """"
   RunPrereq = 1
   Exit Function
  End If

  ' Check If destination folder exists
  If Not oFSO.FolderExists (sDestinationDir) Then
   DebugEcho "Destination folder does not exist: """ & sDestinationDir & """, trying to create it..."
   oFSO.CreateFolder (sDestinationDir)
   If Not oFSO.FolderExists (sDestinationDir) Then
    LogEvent EVENT_ERROR, "Could not create the destination folder: """ & sDestinationDir & """"
    RunPrereq = 2
    Exit Function
   End If
  End If

  ' Copy the file
  oFSO.CopyFile sSource, sDestination, True
  LogEvent EVENT_INFORMATION, "Copied file """ & sSource & """ to """ & sDestination & """"

 End Select

End Function

'--------------------------------------
' Check is the product is installed, and if so, then if the version is same or newer as requested (using text comparison of the version string)
Function IsProductInstalled (ByRef sProductName, ByRef sProductVersion)
 Dim sInstalledVersion

 IsProductInstalled = False

 sInstalledVersion = GetRegProductVer (oReg, REG_UNINSTALL_64, sProductName)     ' x64

 If IsNull (sInstalledVersion) Then
  sInstalledVersion = GetRegProductVer (oReg, REG_UNINSTALL_86, sProductName)    ' x86
 End If

 If IsNull (sInstalledVersion) Then
  DebugEcho "Product """ & sProductName & """ is not installed with version " & sProductVersion
 Else
  IsProductInstalled = True
  If sProductVersion = sInstalledVersion Then
   DebugEcho "Version installed: " & sInstalledVersion & ", same as required: " & sProductVersion
  Else
   If 1 = StrComp (sInstalledVersion, sProductVersion, vbTextCompare) Then
    DebugEcho "Version installed: " & sInstalledVersion & ", newer than required: " & sProductVersion
   Else
    IsProductInstalled = False
    DebugEcho "Version installed: " & sInstalledVersion & ", older than required: " & sProductVersion
   End If
  End If
 End If

End Function

'--------------------------------------
' Copies file, but only if it does not exist at destination yet, or if it is newer on the source
Function CopyFileIfNewer (ByRef sFileName, ByRef sSrcDirPath, ByRef sDstDirPath)
 Dim sSrcFullPath, sDstFullPath

 sSrcFullPath = sSrcDirPath & "\" & sFileName
 sDstFullPath = sDstDirPath & "\" & sFileName

 ' Check If source file exists
 If Not oFSO.FileExists (sSrcFullPath) Then
  DebugEcho "Source file does not exist: """ & sSrcFullPath & """"
  CopyFileIfNewer = False
  Exit Function
 End If

 ' Check If destination folder exists
 If Not oFSO.FolderExists (sDstDirPath) Then
  DebugEcho "Destination folder does not exist: """ & sDstDirPath & """, going to create it"
  oFSO.CreateFolder (sDstDirPath)
  If Not oFSO.FolderExists (sDstDirPath) Then
   LogEvent EVENT_ERROR, "Could not create the destination folder: """ & sDstDirPath & """"
   CopyFileIfNewer = False
   Exit Function
  End If
 End If

 CopyFileIfNewer = True
 ' Check If the source file is newer
 If IsFileNewer (sSrcFullPath, sDstFullPath) Then
  ' Copy the file
  oFSO.CopyFile sSrcFullPath, sDstFullPath, True
  LogEvent EVENT_INFORMATION, "Copied file """ & sSrcFullPath & """ to """ & sDstDirPath & """"
 Else
  LogEvent EVENT_INFORMATION, "No need to copy file """ & sSrcFullPath & """ to """ & sDstDirPath & """"
 End If

End Function

'--------------------------------------
Function IsFileNewer (ByRef sSrcFullPath, ByRef sDstFullPath)
 Dim oSrcFile, oDstFile, dSrcFile, dDstFile

 IsFileNewer = False

 If oFSO.FileExists (sSrcFullPath) Then
  If oFSO.FileExists (sDstFullPath) Then
   Set oSrcFile = oFSO.GetFile (sSrcFullPath)
   Set oDstFile = oFSO.GetFile (sDstFullPath)
   dSrcFile = oSrcFile.DateLastModified
   dDstFile = oDstFile.DateLastModified
   If DateDiff ("d", dSrcFile, dDstFile) > 0 Then
    IsFileNewer = True
   End If
   Set oSrcFile = Nothing
   Set oDstFile = Nothing
  Else
   IsFileNewer = True
  End If
 End If

End Function

'--------------------------------------
Function GetRegProductVer (ByRef oReg, ByRef sRegBranch, ByRef sProductDisplayName)
 Dim aSubKeys, sSubKey, aValues, aTypes
 Dim i, j
 Dim sRegDisplayName, sRegDisplayVersion

 GetRegProductVer = Null

 oReg.EnumKey HKEY_LOCAL_MACHINE, sRegBranch, aSubKeys

 If Not IsArray (aSubKeys) Then
  LogEvent EVENT_ERROR, "Registry does not contain keys under """ & sRegBranch & """"
  Exit Function
 End If

 DebugEcho "Loaded keys under """ & sRegBranch & """"

 For Each sSubKey In aSubKeys
  oReg.EnumValues HKEY_LOCAL_MACHINE, sRegBranch & "\" & sSubKey, aValues, aTypes

  If (8191 < VarType (aValues)) And (8191 < VarType (aTypes)) Then
   i = UBound (aValues)
   For j = 0 To i
    If "DISPLAYNAME" = UCase (aValues(j)) And REG_SZ = aTypes(j) Then
     oReg.GetStringValue HKEY_LOCAL_MACHINE, sRegBranch & "\" & sSubKey, "DisplayName", sRegDisplayName
     If 8 = VarType (sRegDisplayName) Then
      If sProductDisplayName = sRegDisplayName Then
       DebugEcho "Product """ & sProductDisplayName & """ found in Registry"
       oReg.GetStringValue HKEY_LOCAL_MACHINE, sRegBranch & "\" & sSubKey, "DisplayVersion", sRegDisplayVersion
       If 8 = VarType (sRegDisplayVersion) Then
        DebugEcho "Version """ & sRegDisplayVersion & """ found in Registry"
        GetRegProductVer = sRegDisplayVersion
       Else
        DebugEcho "Version is not found in Registry"
        GetRegProductVer = ""
       End If
       Exit For
      End If
     End If
    End If
   Next
  End If
 Next

End Function

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

'--------------------------------------
Function GetDirFromPath (ByRef sFullPath)
 Dim iPos

 iPos = InStrRev (sFullPath, "\")
 If iPos > 1 Then GetDirFromPath = Left (sFullPath, iPos - 1) Else GetDirFromPath = ""

End Function

'--------------------------------------
Function GetFileFromPath (ByRef sFullPath)
 Dim iPos

 iPos = InStrRev (sFullPath, "\")
 If (iPos > 1) And (iPos < (Len (sFullPath))) Then GetFileFromPath = Mid (sFullPath, iPos + 1) Else GetFileFromPath = sFullPath

End Function

'--------------------------------------
Sub DebugEcho (ByVal sString)

 If DEBUG_IT Then WScript.Echo "> " & sString

End Sub

'--------------------------------------
Sub LogEvent (ByVal iEventType, ByVal sString)

 If DEBUG_IT Then WScript.Echo "> " & sString Else oShell.LogEvent iEventType, sEventLog & sString

End Sub

'--------------------------------------
Sub LogEventExit (iLevel, sMessage, iExit)

 If DEBUG_IT Then WScript.Echo aLogLevels(iLevel) & ": " & sMessage Else oShell.LogEvent iLevel, sMessage

 If iExit > 0 Then WScript.Quit (iExit)

End Sub
