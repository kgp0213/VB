VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   3360
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "重新输入"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "请留意此处！"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "请输入欲显示内容"
Else
Label1.Caption = Text1.Text
End If

End Sub

Private Sub Command2_Click()
Text1.Text = ""
End Sub

Private Sub Text1_Change()

End Sub
