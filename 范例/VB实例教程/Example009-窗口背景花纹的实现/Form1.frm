VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "背景花纹"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BackGround(f As Form, pic As PictureBox)
    For i = 0 To (f.ScaleWidth \ pic.Width)
        For j = 0 To (f.ScaleHeight \ pic.Height)
            PaintPicture pic.Picture, i * pic.Width, j * pic.Height
        Next
    Next
'将pic中的图片铺满整个窗口作为窗口背景花纹
End Sub

Private Sub Form_Load()
    Me.AutoRedraw = True
    '窗口可以自动重画
    Pic1.Visible = False
    '隐藏
    Pic1.BorderStyle = 0
    '没有边框
    Pic1.AutoSize = True
    '根据载入的图片自动调节大小
    Pic1.Picture = LoadPicture(App.Path + "\3.bmp")
    '载入图片
    'App指本程序
    'App.Path指本程序路径
    '文件3.bmp存储在本程序目录中
    BackGround Me, Pic1
    '调用子过程BackGround将Pic1中的图片充满窗口
End Sub
Private Sub Form_Resize()
   BackGround Me, Pic1
End Sub
