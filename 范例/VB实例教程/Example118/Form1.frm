VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   4170
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command_Close 
      Caption         =   "Close the Calculator"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   3135
   End
   Begin VB.CommandButton Command_Start 
      Caption         =   "Start the Calculator"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function WaitForSingleObject Lib "kernel32" _
         (ByVal hHandle As Long, _
         ByVal dwMilliseconds As Long) As Long

Private Declare Function FindWindow Lib "user32" _
         Alias "FindWindowA" _
         (ByVal lpClassName As String, _
         ByVal lpWindowName As String) As Long

Private Declare Function PostMessage Lib "user32" _
         Alias "PostMessageA" _
         (ByVal hwnd As Long, _
         ByVal wMsg As Long, _
         ByVal wParam As Long, _
         ByVal lParam As Long) As Long

Private Declare Function IsWindow Lib "user32" _
         (ByVal hwnd As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" _
         (ByVal dwDesiredAccess As Long, _
         ByVal bInheritHandle As Long, _
         ByVal dwProcessId As Long) As Long
      
Private Declare Function GetWindowThreadProcessId Lib "user32" _
         (ByVal hwnd As Long, _
         lpdwProcessId As Long) As Long

'Constants that are used by the API
Const WM_CLOSE = &H10
Const INFINITE = &HFFFFFFFF
Const SYNCHRONIZE = &H100000


Private Sub Command_Close_Click()
      'Closes Windows Calculator
         Dim hWindow As Long
         Dim hThread As Long
         Dim hProcess As Long
         Dim lProcessId As Long
         Dim lngResult As Long
         Dim lngReturnValue As Long

         hWindow = FindWindow(vbNullString, "计算器")
         hThread = GetWindowThreadProcessId(hWindow, lProcessId)
         hProcess = OpenProcess(SYNCHRONIZE, 0&, lProcessId)
         lngReturnValue = PostMessage(hWindow, WM_CLOSE, 0&, 0&)
         lngResult = WaitForSingleObject(hProcess, INFINITE)

         'Does the handle still exist?
         DoEvents
         hWindow = FindWindow(vbNullString, "计算器")
         If IsWindow(hWindow) = 1 Then
            'The handle still exists. Use the TerminateProcess function
            'to close all related processes to this handle.
            MsgBox "Handle still exists."
         Else
            'Handle does not exist.
            MsgBox "All Program Instances Closed."
         End If
End Sub

Private Sub Command_Start_Click()
    'Starts Windows Calculator
    Shell "calc.exe", vbNormalNoFocus
End Sub
