VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "输入和消息"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   4890
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "单击这里！"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim message, title, defaultValue As String
    Dim myValue As String
    message = "Enter a value between 1 and 3"   '设置提示信息
    title = "InputBox Demo"                      '设置标题
    defaultValue = "1"                           '设置默认值
    
    myValue = InputBox(message, title, defaultValue, 100, 100)
   '显示输入对话框
   If myValue = "" Then
        MsgBox "没有输入任何值！", vbInformation + vbOKOnly, "示例"
    Else
        Text1.Text = myValue
    End If
End Sub
