VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mouse"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   3765
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command_RClick 
      Caption         =   "ÓÒ¼üµ¥»÷"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command_LDClick 
      Caption         =   "×ó¼üË«»÷"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command_LClick 
      Caption         =   "×ó¼üµ¥»÷"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command_Move 
      Caption         =   "ÒÆ¶¯Êó±ê"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub mouse_event Lib "user32" _
                (ByVal dwFlags As Long, _
                ByVal dx As Long, _
                ByVal dy As Long, _
                ByVal cButtons As Long, _
                ByVal dwExtraInfo As Long)
Private Declare Function SetCursorPos Lib "user32" _
                (ByVal x As Long, _
                ByVal y As Long) _
                As Long
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Private Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Private Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Private Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Private Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up

Private Sub Command_LClick_Click()
    Call SetCursorPos(10, 10)
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
End Sub

Private Sub Command_LDClick_Click()
    Call SetCursorPos(200, 200)
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
End Sub

Private Sub Command_Move_Click()
    Call mouse_event(MOUSEEVENTF_MOVE, 100, 100, 0, 0)
End Sub

Private Sub Command_RClick_Click()
    Call SetCursorPos(200, 200)
    Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
    Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
End Sub
