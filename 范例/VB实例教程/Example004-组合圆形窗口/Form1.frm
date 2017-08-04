VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5880
   ScaleWidth      =   5880
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Click()
'单击窗口则调用SetWindow子过程设置窗口形状
    SetWindow Form1
    '设置窗口形状
End Sub

Private Sub Form_DblClick()
'双击窗口则调用Reset子过程恢复窗口形状
    Reset Form1
    '将窗口恢复为矩形
End Sub
