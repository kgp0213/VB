VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "图片按钮"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4305
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check1 
      Caption         =   "无效"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      DisabledPicture =   "Form1.frx":0000
      Enabled         =   0   'False
      Height          =   2295
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
'
    If Check1.Value = 1 Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Command1.Picture = LoadPicture(App.Path + "\1.bmp")
    Command1.DownPicture = LoadPicture(App.Path + "\2.bmp")
    Command1.DisabledPicture = LoadPicture(App.Path + "\3.bmp")
End Sub
