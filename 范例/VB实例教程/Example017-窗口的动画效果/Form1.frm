VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6180
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function AnimateWindow _
Lib "user32" _
( _
    ByVal hWnd As Long, _
    ByVal dwTime As Long, _
    ByVal dwFlags As Long _
 ) As Long
Private Const AW_HOR_POSITIVE = &H1
Private Const AW_HOR_NEGATIVE = &H2
Private Const AW_VER_POSITIVE = &H4
Private Const AW_VER_NEGATIVE = &H8
Private Const AW_CENTER = &H10
Private Const AW_HIDE = &H10000
Private Const AW_ACTIVATE = &H20000
Private Const AW_SLIDE = &H40000
Private Const AW_BLEND = &H80000
Private Sub Form_Load()
    AnimateWindow Me.hWnd, 10000, AW_BLEND Or AW_ACTIVATE
    '还可以使用前面定义的其他常量显示不同的动画效果
    Form1.Refresh
End Sub

