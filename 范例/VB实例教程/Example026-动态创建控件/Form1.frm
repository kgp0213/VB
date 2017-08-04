VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "动态创建控件"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "删除控件"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加控件"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents NewButton As CommandButton
Attribute NewButton.VB_VarHelpID = -1
'通过使用WithEvents关键字声明一个对象变量为新的命令按钮

Private Sub Command1_Click()
If NewButton Is Nothing Then
    Set NewButton = Controls.Add("VB.CommandButton", "cmdNew", Form1)
    '增加新的按钮cmdNew
    NewButton.Move Command1.Left + Command1.Width + 240, Command1.Top
    '确定新增按钮cmdNew的位置
    NewButton.Caption = "动态添加的按钮"
    NewButton.Visible = True
    '显示该按钮
End If
End Sub

Private Sub Command2_Click()
If NewButton Is Nothing Then
    Exit Sub
Else
   Controls.Remove NewButton
   Set NewButton = Nothing
   End If
End Sub
Private Sub NewButton_click()
    MsgBox "这是动态增加的按钮，你可以单击“删除控件”按钮删除它", vbDefaultButton1, "Click"
End Sub

