VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateMutex Lib "kernel32" _
                Alias "CreateMutexA" _
                (lpMutexAttributes As SECURITY_ATTRIBUTES, _
                ByVal bInitialOwner As Long, _
                ByVal lpName As String) _
                As Long

Private Declare Function ReleaseMutex Lib "kernel32" _
                (ByVal hMutex As Long) _
                As Long
                
Private Declare Function CloseHandle Lib "kernel32" _
                (ByVal hObject As Long) _
                As Long
                
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Const ERROR_ALREADY_EXISTS = 183&

Private mlngHMutex As Long

Private Sub Form_Load()
    If IsPrevAppRunning = True Then
        MsgBox "This app is already running.", _
                vbOKOnly + vbExclamation, "App Already In Use"
        Unload Me
        Set frmMutex = Nothing
    End If
End Sub

Private Sub Form_Terminate()
    If mlngHMutex <> 0 Then
    'Close the mutex.
        ReleaseMutex mlngHMutex
        CloseHandle mlngHMutex
    End If
End Sub

Private Function IsPrevAppRunning() As Boolean
    On Error GoTo error_IsPrevAppRunning
    
    Dim lngVBRet As Long
    Dim sa As SECURITY_ATTRIBUTES
    mlngHMutex = CreateMutex(sa, True, "Test")
    lngVBRet = Err.LastDllError
    If lngVBRet = ERROR_ALREADY_EXISTS Then
    'This app is already running.
        IsPrevAppRunning = True
    End If
    Exit Function
    
error_IsPrevAppRunning:
    IsPrevAppRunning = False
End Function
