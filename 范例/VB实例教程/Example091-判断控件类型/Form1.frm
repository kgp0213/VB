VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   2850
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "列举窗体中控件的类型"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
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

Private Sub Command1_Click()
   Dim cn As Control
     For Each cn In Form1
        Select Case TypeName(cn)
             Case "CommandButton"
                MsgBox "这是一个按钮，名称为：" + cn.Name
             Case "PictureBox"
                MsgBox "这是一个图片框，名称为：" + cn.Name
             Case "Label"
                MsgBox "这是一个标签，名称为：" + cn.Name
             Case "TextBox"
                MsgBox "这是一个文本框，名称为：" + cn.Name
             Case "CheckBox"
                MsgBox "这是一个复选框，名称为：" + cn.Name
         End Select
     Next

End Sub
