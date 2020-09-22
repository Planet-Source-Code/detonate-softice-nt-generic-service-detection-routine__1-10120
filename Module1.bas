Attribute VB_Name = "Module1"
'// Numega SoftICE-NT Detection Routine for WinNT/2K
'// Developed for Planet Source Code (www.planetsourcecode.com/vb)
'// by Detonate (detonate@start.com.au), July 2000
'// You are free to use, modify, publish and abuse this code at will,
'// as long as this header remains intact and unmodified.
'// Greets to Joox (SoftICE-9x Detection) and Kevin Lingofelter (Generic service routines)

'// Joox's SoftICE-9x detection routine can be found at:
'   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=7600

'// Kevin has modified Joox's routine to also detect some SoftICE-NT variants (detects his 4.x, but not my 3.5x)
'   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=10000

'// Now that this has been published, it will be easy for crhackers™ to figure out how it is done,
'// but on the other hand, it's not every day that SoftICE veterans sift through Visual Basic source :-)
'// Either way, if we've wasted just 1 minute of the crhackers™ time, then our job is done.

Public Type SERVICE_STATUS
   dwServiceType              As Long
   dwCurrentState             As Long
   dwControlsAccepted         As Long
   dwWin32ExitCode            As Long
   dwServiceSpecificExitCode  As Long
   dwCheckPoint               As Long
   dwWaitHint                 As Long
End Type

Public Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Public Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Public Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SC_MANAGER_CONNECT = &H1
Public Const SC_MANAGER_CREATE_SERVICE = &H2
Public Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Public Const SC_MANAGER_LOCK = &H8
Public Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Public Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
Public Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
                                      SC_MANAGER_CONNECT Or _
                                      SC_MANAGER_CREATE_SERVICE Or _
                                      SC_MANAGER_ENUMERATE_SERVICE Or _
                                      SC_MANAGER_LOCK Or _
                                      SC_MANAGER_QUERY_LOCK_STATUS Or _
                                      SC_MANAGER_MODIFY_BOOT_CONFIG)

Public Function TestIce(xService As String) As Integer
On Error Resume Next
Dim ptypServiceStatus As SERVICE_STATUS
   plngServerHwnd = OpenSCManager(mstrComputerName, vbNullString, SC_MANAGER_ALL_ACCESS)
   If plngServerHwnd = 0 Then
      TestIce = 0
      Exit Function
   Else
      plngServiceHwnd = OpenService(plngServerHwnd, xService, plngAccessType)
      If plngServiceHwnd = 0 Then
         TestIce = 1
         Call CloseServiceHandle(plngServerHwnd)
         Exit Function
      Else
         Call CloseServiceHandle(plngServiceHwnd)
         Call CloseServiceHandle(plngServerHwnd)
         TestIce = 2
         Exit Function
      End If
   End If
End Function
