VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "格式转换"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保存文件"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开文件"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    CommonDialog1.DialogTitle = "打开文件"
    CommonDialog1.Filter = "所有支持的格式" + _
                            "(*.bmp;*.jpg;*.gif;*.pcx;*.ico)|" + _
                            "*.bmp;*.jpg;*.gif;*.pcx;*.ico)"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        If Err <> 32755 Then
           Dim OpenFileName As String
            OpenFileName = CommonDialog1.FileName
            Picture1.Picture = LoadPicture(OpenFileName)
        End If
    End If
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    CommonDialog1.DialogTitle = "保存为BMP文件"
    CommonDialog1.Filter = "位图文件(*.bmp)|*.bmp"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        If Err <> 32755 Then
           Dim SaveBmpName As String
            SaveBmpName = CommonDialog1.FileName
            SavePicture Picture1.Image, SaveBmpName
        End If
    End If
End Sub

Private Sub Form_Load()
    CommonDialog1.CancelError = True
End Sub
