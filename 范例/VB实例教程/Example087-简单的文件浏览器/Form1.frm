VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "简单的文件浏览器"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   6180
   StartUpPosition =   3  '窗口缺省
   Begin VB.FileListBox File1 
      Height          =   4230
      Left            =   2160
      Pattern         =   "*.exe"
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Height          =   3660
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
'当目录发生改变时使File1也随之改变
'以便显示该目录中的文件
    File1.Path = Me.Dir1.Path
End Sub

Private Sub Drive1_Change()
'当驱动器发生改变时使Dir1与其保持一致
On Error GoTo IFerr    '拦截错误
Dir1.Path = Drive1.Drive
Exit Sub
IFerr:                 '如果磁盘错误
    MsgBox "请确认驱动器是否准备好或者磁盘已经不可用!", _
            vbOKOnly + vbExclamation
    Drive1.Drive = Dir1.Path  '忽略驱动器改变
    Exit Sub
End Sub


