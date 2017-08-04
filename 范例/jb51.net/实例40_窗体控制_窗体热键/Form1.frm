VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "窗体热键"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   2745
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "该窗体热键为Ctrl+Alt+Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'热键为
'Ctl+Alt+Z.

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_SETHOTKEY = &H32
Private Const HOTKEYF_SHIFT = &H1
Private Const HOTKEYF_CONTROL = &H2
Private Const HOTKEYF_ALT = &H4

Private Sub Form_Load()
   Dim l As Long
   Dim wHotkey As Long
   
   wHotkey = (HOTKEYF_ALT Or HOTKEYF_CONTROL) * (2 ^ 8) + 90
   l = SendMessage(Me.hwnd, WM_SETHOTKEY, wHotkey, 0)
End Sub
