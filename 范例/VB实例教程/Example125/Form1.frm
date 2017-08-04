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
Private Declare Function SendMessage Lib "user32" _
                Alias "SendMessageA" _
                (ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) _
                As Long

Const WM_SETHOTKEY = &H32
Const HOTKEYF_ALT = &H4

Private Sub Form_Load()
    Dim wHotkey As Integer
    wHotkey = (HOTKEYF_ALT) * &H100 + Asc("B")
    Call SendMessage(Me.hwnd, WM_SETHOTKEY, wHotkey, 0)
End Sub
