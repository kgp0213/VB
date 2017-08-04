VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "反转颜色"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   5745
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   2880
      ScaleHeight     =   3075
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3120
      Left            =   120
      ScaleHeight     =   3060
      ScaleWidth      =   2565
      TabIndex        =   0
      Top             =   360
      Width           =   2625
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Picture1.Picture = LoadPicture(App.Path + "\鸟.bmp")
End Sub

Private Sub Picture2_Click()
   Picture2.Cls
   Picture2.FillColor = &HFFFFF
   Picture2.PaintPicture Picture1.Picture, 0, 0, , , 0, 0, , , &H660046
   '以下代码采用逐象素方式反转颜色
   'Dim i, j As Integer
   'Dim c1 As Long
   'Dim r, g, b As Integer
   'Form1.ScaleMode = 3
   'Picture1.ScaleMode = 3
   'Picture2.ScaleMode = 3
   'For i = 0 To Picture1.Width
   'For j = 0 To Picture1.Height
    'c1 = Picture1.Point(i, j)
    'r = c1 And &HFF
    'g = (c1 And 62580) / 256
    'b = (c1 And &HFF0000) / 65536
    'r = 255 - r
    'g = 255 - g
    'b = 255 - b
    'If r < 0 Then r = 0
    'If r > 255 Then r = 255
    'If g < 0 Then g = 0
    'If g > 255 Then g = 255
    'If b < 0 Then b = 0
    'If b > 255 Then b = 255
    'Picture2.PSet (i, j), RGB(r, g, b)
    'Next
    'Next
 End Sub
