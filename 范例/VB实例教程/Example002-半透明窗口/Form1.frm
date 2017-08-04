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
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
                (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) _
                As Long
Private Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
                (ByVal hwnd As Long, _
                ByVal nIndex As Long) _
                As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA As Long = &H2
Private Const WS_EX_LAYERED As Long = &H80000

Private Declare Function SetLayeredWindowAttributes Lib "user32" _
                (ByVal hwnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Long, _
                ByVal dwFlags As Long) _
                As Long

Private Sub Form_Load()
  Dim p As Long
  p = GetWindowLong(Me.hwnd, GWL_EXSTYLE) '取得当前窗口属性
  Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, p Or WS_EX_LAYERED)
  '加上一个透明属性
  Call SetLayeredWindowAttributes(Me.hwnd, 0, 0, LWA_ALPHA)
End Sub
