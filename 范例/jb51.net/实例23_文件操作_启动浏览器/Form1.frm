VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":030A
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK,  to http://www.edu.cn"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2520
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call Shell("MyBat.bat", vbMinimizedFocus)
End Sub
'调用浏览器的简单方法：运行SHELL命令执行一个批文件 ( .bat ) 。

