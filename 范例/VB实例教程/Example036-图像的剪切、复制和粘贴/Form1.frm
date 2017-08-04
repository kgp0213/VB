VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "picclp32.ocx"
Begin VB.Form Form1 
   Caption         =   "复制剪切和粘贴"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   7575
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture2 
      Height          =   3975
      Left            =   4200
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   4
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton CmdPaste 
      Caption         =   "粘贴"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton CmdCut 
      Caption         =   "剪切"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "复制"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.Shape Shape1 
         BorderStyle     =   3  'Dot
         Height          =   1455
         Left            =   960
         Top             =   1080
         Width           =   1575
      End
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   0
      Top             =   120
      _ExtentX        =   7011
      _ExtentY        =   7646
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag1 As Boolean
Private Sub Form_Load()
    Shape1.Visible = False
    Shape1.BorderStyle = 3
    flag1 = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, Y As Single)
'开始选择区域
    Shape1.Left = X
    Shape1.Top = Y
    flag1 = True
    '设置标志变量并将Shape1的左上角移动到鼠标所在点
End Sub

Private Sub Picture1_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, Y As Single)
'在选定区域过程中随着鼠标移动产生虚线框
   If Button = 1 Then
       If flag1 = True Then
       '如果是处在正在选择区域状态
            Shape1.Width = Abs(X - Shape1.Left)
            Shape1.Height = Abs(Y - Shape1.Top)
            Shape1.Visible = True
            Picture1.Refresh
        Else
            Shape1.Visible = False
        End If
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             X As Single, Y As Single)
    flag1 = False
    '结束选择区域状态
End Sub
Private Sub CmdCopy_Click()
'通过PictureClip控件作为中间对象将Picture1中由Shape1表明的图像块
'复制到剪贴板上
    If Shape1.Visible = True Then
    '如果有选定的图像块
        Clipboard.Clear    '清空剪贴扳
        On Error Resume Next
        PictureClip1.Picture = Picture1.Picture
        PictureClip1.ClipX = Shape1.Left
        PictureClip1.ClipY = Shape1.Top
        PictureClip1.ClipWidth = Shape1.Width
        PictureClip1.ClipHeight = Shape1.Height
        Clipboard.SetData PictureClip1.Clip
    End If
End Sub

Private Sub CmdCut_Click()
Const vbMergePaint = &HBB0226
    If Shape1.Visible = True Then
        Clipboard.Clear    '清空剪贴扳
        On Error Resume Next
        PictureClip1.Picture = Picture1.Picture
        PictureClip1.ClipX = Shape1.Left
        PictureClip1.ClipY = Shape1.Top
        PictureClip1.ClipWidth = Shape1.Width
        PictureClip1.ClipHeight = Shape1.Height
        Clipboard.SetData PictureClip1.Clip
        '如果有选定的图像块则复制到剪贴板
    
        Picture1.PaintPicture Picture1.Picture, _
             Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, _
             Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, _
             vbMergePait
        '使用OR运算使Picture1中Shape1所标识的部分清空
        
    End If
End Sub

Private Sub CmdPaste_Click()
'粘贴
    Picture2.Picture = Clipboard.GetData
End Sub


