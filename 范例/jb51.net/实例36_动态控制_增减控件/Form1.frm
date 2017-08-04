VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "动态增减控件的例子"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "删除控件"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "增加控件"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '通过使用WithEvents关键字声明一个对象变量为新的命令按钮
    Private WithEvents NewButton As CommandButton
Attribute NewButton.VB_VarHelpID = -1
    '增加控件
    Private Sub Command1_Click()
    If NewButton Is Nothing Then
    '增加新的按钮cmdNew
    Set NewButton = Controls.Add("VB.CommandButton", "cmdNew", Me)
    '确定新增按钮cmdNew的位置
    NewButton.Move Command1.Left + Command1.Width + 240, Command1.Top
    NewButton.Caption = "新增按钮"
    NewButton.Visible = True
    End If
    End Sub
    '删除控件(注：只能删除动态增加的控件)
    Private Sub Command2_Click()
    If NewButton Is Nothing Then
    Else
    Controls.Remove NewButton
    Set NewButton = Nothing
    End If
    End Sub
    '新增控件的单击事件
    Private Sub NewButton_Click()
     MsgBox "您选中的是动态增加的按钮!"
    End Sub

