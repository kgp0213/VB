VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "图像的合成"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   569
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   619
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdCls 
      Caption         =   "清除"
      Height          =   495
      Left            =   1200
      TabIndex        =   12
      Top             =   6840
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CmnDlg1 
      Left            =   3480
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "保存"
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Cmdunite 
      Caption         =   "合成"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   6840
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   3375
      Left            =   3720
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   9
      Top             =   4920
      Width           =   5295
   End
   Begin MSComctlLib.Slider SliderY 
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   6240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
   End
   Begin MSComctlLib.Slider SliderX 
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   5880
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "合成方式"
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   2655
      Begin VB.OptionButton Option3 
         Caption         =   "覆盖"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "异或"
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "与"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "选择背景"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton CmdFore 
      Caption         =   "选择前景"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   3840
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   341
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   3120
      Left            =   120
      ScaleHeight     =   204
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   211
      TabIndex        =   0
      Top             =   360
      Width           =   3225
   End
   Begin VB.Label Label2 
      Caption         =   "Y轴"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "X轴"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   5880
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PLeft As Integer
Dim PTop As Integer
Dim Mode As Integer
Private Sub Form_Load()
    SliderX.Max = Picture1.Width
    SliderY.Max = Picture1.Height
    Mode = 1
End Sub
Private Sub CmdFore_Click()
'打开前景图像
   On Error GoTo Err_handle
   CmnDlg1.DialogTitle = "打开"
   CmnDlg1.ShowOpen
   Picture1.Picture = LoadPicture(CmnDlg1.FileName)
Err_handle:   Exit Sub
End Sub

Private Sub CmdBack_Click()
'打开背景图像
   On Error GoTo Err_handle
   CmnDlg1.DialogTitle = "打开"
   CmnDlg1.ShowOpen
   Picture2.Picture = LoadPicture(CmnDlg1.FileName)
Err_handle:   Exit Sub
End Sub

Private Sub CmdCls_Click()
'清除合成图像中的前景图像
    Picture3.Cls
End Sub
Private Sub Cmdunite_Click()
'合并前景图像和背景图像
    On Error Resume Next
    Picture3.Width = Picture2.Width
    Picture3.Height = Picture2.Height
    '合成后的图像和背景图保持同样大小
    Picture3.Picture = Picture2.Picture
    If Mode = 1 Then
        Picture3.PaintPicture Picture1.Picture, PLeft, PTop, , , , , , , &H8800C6
    ElseIf Mode = 2 Then
        Picture3.PaintPicture Picture1.Picture, PLeft, PTop, , , , , , , &H660046
   ElseIf Mode = 3 Then
        Picture3.PaintPicture Picture1.Picture, PLeft, PTop
   End If
End Sub
Private Sub CmdSave_Click()
'保存合成的图像
   On Error GoTo Err_handle
   CmnDlg1.DialogTitle = "保存"
   CmnDlg1.Filter = ("位图文件(*.bmp)|*.bmp")
   CmnDlg1.ShowSave
   SavePicture Picture3.Image, CmnDlg1.FileName
Err_handle:   Exit Sub
End Sub
'设置合并方式
Private Sub Option1_Click()
    If Option1.Value = True Then Mode = 1
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then Mode = 2
End Sub

Private Sub Option3_Click()
    If Option3.Value = True Then Mode = 3
End Sub

'设置X轴和Y轴方向的偏移量
Private Sub SliderX_Scroll()
    PLeft = SliderX.Value
End Sub

Private Sub SliderY_Scroll()
    PTop = SliderY.Value
End Sub
