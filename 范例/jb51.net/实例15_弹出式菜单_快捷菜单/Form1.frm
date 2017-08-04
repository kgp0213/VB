VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1590
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   2850
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "单击鼠标右键，可以使用快捷菜单！"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Menu Filemenu 
      Caption         =   "文件(&F)"
      Begin VB.Menu NewMenu 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu OpenMenu 
         Caption         =   "打开(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu SaveMenu 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu Menu1 
         Caption         =   "-"
      End
      Begin VB.Menu CloseMenu 
         Caption         =   "关闭"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "编辑(&E)"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu Filemenu
End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu Filemenu
End If
End Sub
