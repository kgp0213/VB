VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "锐化"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   383
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "锐化"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   2880
      ScaleHeight     =   205
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3120
      Left            =   120
      ScaleHeight     =   204
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   0
      Top             =   360
      Width           =   2650
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim r2, g2, b2 As Integer
    Dim r1, g1, b1 As Integer
    Dim c1 As Long
    Dim c2 As Long
    Dim x0 As Integer
    Dim y0 As Integer
    Screen.MousePointer = 11
    For x0 = 1 To Picture1.Width - 2
    For y0 = 1 To Picture2.Height - 2
        c1 = Picture1.Point(x0, y0)
        r1 = (c1 And &HFF)
        g1 = (c1 And 62580) / 256
        b1 = (c1 And &HFF0000) / 65536
        '获得Picture1中指定点(x0,y0)的R、G、B分量值
        
        c2 = Picture1.Point(x0 - 1, y0 - 1)
        r2 = (c2 And &HFF)
        g2 = (c2 And 62580) / 256
        b2 = (c2 And &HFF0000) / 65536
        '获得Picture1中与(x0,y0)点相邻的点的R、G、B分量值
        
        r1 = r1 + (r1 - r2) \ 2
        g1 = g1 + (g1 - g2) \ 2
        b1 = b1 + (b1 - b2) \ 2
        '计算锐化后的R、G、B分量值
        If r1 > 255 Then r1 = 255
        If r1 < 0 Then r1 = 0
        If b1 > 255 Then b1 = 255
        If b1 < 0 Then b1 = 0
        If g1 > 255 Then g1 = 255
        If g1 < 0 Then g1 = 0
        Picture2.PSet (x0, y0), RGB(r1, g1, b1)
        '画出锐化后的(x0,y0)
        DoEvents
    Next
    Next
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Picture1.Picture = LoadPicture(App.Path + "\鸟.bmp")
End Sub

