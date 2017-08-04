VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "缩小"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   2415
      Left            =   3840
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "保存文件"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton CmdScale 
      Caption         =   "缩小"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "打开文件"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "显示比例"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   4200
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
         Caption         =   "25%"
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "50%"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   855
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
Dim i, j As Integer
Dim r, g, b As Integer
Dim r1, g1, b1 As Integer
Dim r2, g2, b2 As Integer
Dim r3, g3, b3 As Integer
Dim r4, g4, b4 As Integer
Dim c1, c2, c3, c4 As Long
If flag = 2 Then
'将图像缩小为原来的四分之一
    Picture2.Width = Picture1.Width / 2
    Picture2.Height = Picture1.Height / 2
    For i = 0 To Picture2.Width Step 1
    For j = 0 To Picture2.Height Step 1
        c1 = Picture1.Point(2 * i, 2 * j)
        r1 = c1 And &HFF
        g1 = (c1 And 62580) / 256
        b1 = (c1 And &HFF0000) / 65536
        
        c2 = Picture1.Point(2 * i, 2 * j + 1)
        r2 = c2 And &HFF
        g2 = (c2 And 62580) / 256
        b2 = (c2 And &HFF0000) / 65536
        
        c3 = Picture1.Point(2 * i + 1, 2 * j)
        r3 = c3 And &HFF
        g3 = (c3 And 62580) / 256
        b3 = (c3 And &HFF0000) / 65536
        
        c4 = Picture1.Point(2 * i + 1, 2 * j + 1)
        r4 = c4 And &HFF
        g4 = (c4 And 62580) / 256
        b4 = (c4 And &HFF0000) / 65536
        r = (r1 + r2 + r3 + r4) / 4
        g = (g1 + g2 + g3 + g4) / 4
        b = (b1 + b2 + b3 + b4) / 4
        Picture2.PSet (i, j), RGB(r, g, b)
    Next
    Next
ElseIf flag = 4 Then
'将图像缩小为原来的十六分之一
    Picture2.Width = Picture1.Width / 4
    Picture2.Height = Picture1.Height / 4
    For i = 0 To Picture1.Width Step 4
    For j = 0 To Picture1.Height Step 4
        c1 = Picture1.Point(i, j)
        Picture2.PSet (i / 4, j / 4), c1
    Next
    Next
ElseIf flag = 3 Then
'利用PictureBox控件缩小图像
    Dim temp As String
    temp = InputBox("请输入倍数", "自定义", 0.5)
    If temp <> "" And temp > 0 And temp < 1 Then
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
