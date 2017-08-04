VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Windows"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   2940
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdShutdown 
      Caption         =   "关机"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdReboot 
      Caption         =   "重新启动"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdLogoff 
      Caption         =   "注销"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const EWX_LogOff As Long = 0
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_REBOOT As Long = 2
Private Const EWX_FORCE As Long = 4
Private Const EWX_POWEROFF As Long = 8

'The ExitWindowsEx function either logs off, shuts down, or shuts
'down and restarts the system.
Private Declare Function ExitWindowsEx Lib "user32" _
                (ByVal dwOptions As Long, _
                ByVal dwReserved As Long) As Long

'The GetLastError function returns the calling thread's last-error
'code value. The last-error code is maintained on a per-thread basis.
'Multiple threads do not overwrite each other's last-error code.
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Type LUID
   UsedPart As Long
   IgnoredForNowHigh32BitPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   TheLuid As LUID
   Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   TheLuid As LUID
   Attributes As Long
End Type

'The GetCurrentProcess function returns a pseudohandle for the
'current process.
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'The OpenProcessToken function opens the access token associated with
'a process.
Private Declare Function OpenProcessToken Lib "advapi32" _
   (ByVal ProcessHandle As Long, _
    ByVal DesiredAccess As Long, _
    TokenHandle As Long) As Long

'The LookupPrivilegeValue function retrieves the locally unique
'identifier (LUID) used on a specified system to locally represent
'the specified privilege name.
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
   Alias "LookupPrivilegeValueA" _
   (ByVal lpSystemName As String, _
    ByVal lpName As String, _
    lpLuid As LUID) As Long

'The AdjustTokenPrivileges function enables or disables privileges
'in the specified access token. Enabling or disabling privileges
'in an access token requires TOKEN_ADJUST_PRIVILEGES access.
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
   (ByVal TokenHandle As Long, _
    ByVal DisableAllPrivileges As Long, _
    NewState As TOKEN_PRIVILEGES, _
    ByVal BufferLength As Long, _
    PreviousState As TOKEN_PRIVILEGES, _
    ReturnLength As Long) As Long

Private Declare Sub SetLastError Lib "kernel32" _
   (ByVal dwErrCode As Long)

Private Const mlngWindows95 = 0
Private Const mlngWindowsNT = 1

Public glngWhichWindows32 As Long

'The GetVersion function returns the operating system in use.
Private Declare Function GetVersion Lib "kernel32" () As Long

Private Sub AdjustToken()
'********************************************************************
'* This procedure sets the proper privileges to allow a log off or a
'* shut down to occur under Windows NT.
'********************************************************************

   Const TOKEN_ADJUST_PRIVILEGES = &H20
   Const TOKEN_QUERY = &H8
   Const SE_PRIVILEGE_ENABLED = &H2

   Dim hdlProcessHandle As Long
   Dim hdlTokenHandle As Long
   Dim tmpLuid As LUID
   Dim tkp As TOKEN_PRIVILEGES
   Dim tkpNewButIgnored As TOKEN_PRIVILEGES
   Dim lBufferNeeded As Long

   'Set the error code of the last thread to zero using the
   'SetLast Error function. Do this so that the GetLastError
   'function does not return a value other than zero for no
   'apparent reason.
   SetLastError 0

   'Use the GetCurrentProcess function to set the hdlProcessHandle
   'variable.
   hdlProcessHandle = GetCurrentProcess()
   OpenProcessToken hdlProcessHandle, _
            (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), _
            hdlTokenHandle

   'Get the LUID for shutdown privilege
   LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid

   tkp.PrivilegeCount = 1    ' One privilege to set
   tkp.TheLuid = tmpLuid
   tkp.Attributes = SE_PRIVILEGE_ENABLED

   'Enable the shutdown privilege in the access token of this process
   AdjustTokenPrivileges hdlTokenHandle, _
            False, _
            tkp, _
            Len(tkpNewButIgnored), _
            tkpNewButIgnored, _
            lBufferNeeded
End Sub

Private Sub cmdLogoff_Click()
   If glngWhichWindows32 = mlngWindowsNT Then
        AdjustToken
   End If
   
   ExitWindowsEx EWX_LogOff, 0
End Sub

Private Sub cmdShutdown_Click()
   If glngWhichWindows32 = mlngWindowsNT Then
        AdjustToken
   End If

   ExitWindowsEx (EWX_SHUTDOWN Or EWX_FORCE Or EWX_POWEROFF), 0
End Sub

Private Sub cmdReboot_Click()
   If glngWhichWindows32 = mlngWindowsNT Then
        AdjustToken
   End If

   ExitWindowsEx EWX_REBOOT, 0
End Sub

Private Sub Form_Load()
'********************************************************************
'* When the project starts, check the operating system used by
'* calling the GetVersion function.
'********************************************************************
    Dim lngVersion As Long

    lngVersion = GetVersion()

    If ((lngVersion And &H80000000) = 0) Then
        glngWhichWindows32 = mlngWindowsNT
    Else
        glngWhichWindows32 = mlngWindows95
    End If
End Sub
