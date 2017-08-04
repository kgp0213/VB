VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "放大"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   608
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   6615
      Left            =   3840
      ScaleHeight     =   437
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   7
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "保存文件"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton CmdScale 
      Caption         =   "放大"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "打开文件"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "显示比例"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   6720
      Width           =   2895
      Begin VB.OptionButton Option3 
         Caption         =   "自定义"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "400%"
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "200%"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3600
      Left            =   120
      ScaleHeight     =   236
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   211
      TabIndex        =   0
      Top             =   120
      Width           =   3225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer

Private Sub CmdOpen_Click()
'打开文件
On Error GoTo Error_Handle
    CommonDialog1.DialogTitle = "打开文件"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        If Err <> 32755 Then
           Dim OpenFileName As String
            OpenFileName = CommonDialog1.FileName
            Picture1.Picture = LoadPicture(OpenFileName)
        End If
    End If
    Exit Sub
Error_Handle: MsgBox Err.Description, vbOKOnly
              Exit Sub
End Sub

Private Sub CmdSave_Click()
On Error GoTo Error_Handle
    CommonDialog1.DialogTitle = "保存为BMP文件"
    CommonDialog1.Filter = "位图文件(*.bmp)|*.bmp"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        If Err <> 32755 Then
           Dim SaveBmpName As String
            SaveBmpName = CommonDialog1.FileName
            SavePicture Picture2.Image, SaveBmpName
        End If
    End If
Error_Handle: MsgBox Err.Description, vbOKOnly
              Exit Sub
End Sub

Private Sub CmdScale_Click()
If flag = 2 Then
'将图像放大2倍
    Picture2.Width = Picture1.Width * 2
    Picture2.Height = Picture1.Height * 2
    For i = 0 To Picture2.Width * 2 - 1 Step 2
    For j = 0 To Picture2.Height * 2 - 1 Step 2
        c = Picture1.Point(i / 2, j / 2)
        Picture2.PSet (i, j), c
        Picture2.PSet (i + 1, j), c
        Picture2.PSet (i, j + 1), c
        Picture2.PSet (i + 1, j + 1), c
    Next
    Next
ElseIf flag = 4 Then
'将图像放大4倍
    Picture2.Width = Picture1.Width * 4
    Picture2.Height = Picture1.Height * 4
    For i = 0 To Picture2.Width * 4 - 3 Step 4
    For j = 0 To Picture2.Height * 4 - 3 Step 4
        c = Picture1.Point(i / 4, j / 4)
        Picture2.PSet (i, j), c
        Picture2.PSet (i, j + 1), c
        Picture2.PSet (i, j + 2), c
        Picture2.PSet (i, j + 3), c
        Picture2.PSet (i + 1, j), c
        Picture2.PSet (i + 1, j + 1), c
        Picture2.PSet (i + 1, j + 2), c
        Picture2.PSet (i + 1, j + 3), c
        Picture2.PSet (i + 2, j), c
        Picture2.PSet (i + 2, j + 1), c
        Picture2.PSet (i + 2, j + 2), c
        Picture2.PSet (i + 2, j + 3), c
        Picture2.PSet (i + 3, j), c
        Picture2.PSet (i + 3, j + 1), c
        Picture2.PSet (i + 3, j + 2), c
        Picture2.PSet (i + 3, j + 3), c
    Next
    Next
ElseIf flag = 3 Then
'利用PictureBox控件放大图像
    temp = InputBox("请输入放大倍数", "自定义", 1.5)
    If temp <> "" And temp > 0 Then
        Picture2.Width = Picture1.Width * temp
        Picture2.Height = Picture1.Height * temp
        Picture2.PaintPicture Picture1.Picture, 0, 0, Picture2.Width, Picture2.Height, _
                              0, 0, Picture1.Width, Picture1.Height
    Else
        MsgBox "输入倍数不符合要求", vbExclamation, "错误"
    End If
End If

End Sub

Private Sub Form_Load()
'为Picture1添加图像并初始化flag变量
    Picture1.Picture = LoadPicture(App.Path + "\鸟.bmp")
    flag = 2
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then flag = 2
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then flag = 4
End Sub

Private Sub Option3_Click()
    If Option3.Value = True Then flag = 3
End Sub
