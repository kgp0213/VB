VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
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
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3120
      Left            =   120
      ScaleHeight     =   204
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   171
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

Private Sub Command1_Click()
    Dim red As Integer
    Dim green As Integer
    Dim blue As Integer
    Dim c1, c2 As Long
    Dim i, j As Integer
    For i = 1 To Picture1.Width - 2 Step 4
        For j = 1 To Picture1.Height - 2 Step 4
            c1 = Picture1.Point(i, j)
            red = c1 And &HFF
            green = (c1 And 62580) / 256
            blue = (c1 And &HFF0000) / 65536
             '颜色处理
            Picture2.PSet (i, j), RGB(red, green, blue)
            Picture2.PSet (i, j + 1), RGB(red, green, blue)
            Picture2.PSet (i, j + 2), RGB(red, green, blue)
            Picture2.PSet (i, j + 3), RGB(red, green, blue)
            
            Picture2.PSet (i + 1, j), RGB(red, green, blue)
            Picture2.PSet (i + 1, j + 1), RGB(red, green, blue)
            Picture2.PSet (i + 1, j + 2), RGB(red, green, blue)
            Picture2.PSet (i + 1, j + 3), RGB(red, green, blue)
        
            Picture2.PSet (i + 2, j), RGB(red, green, blue)
            Picture2.PSet (i + 2, j + 1), RGB(red, green, blue)
            Picture2.PSet (i + 2, j + 2), RGB(red, green, blue)
            Picture2.PSet (i + 2, j + 3), RGB(red, green, blue)
            
            Picture2.PSet (i + 3, j), RGB(red, green, blue)
            Picture2.PSet (i + 3, j + 1), RGB(red, green, blue)
            Picture2.PSet (i + 3, j + 2), RGB(red, green, blue)
            Picture2.PSet (i + 3, j + 3), RGB(red, green, blue)
        Next
    Next

End Sub

Private Sub Form_Load()
    With Form1
        .Caption = "降低分辨率"
        .Height = 4800
        .Left = 0
        .Top = 0
        .Width = 6000
        .ScaleMode = 3          'Pixel
    End With
    With Command1
        .Caption = "降低分辨率"
        .Height = 40
        .Left = 20
        .Top = 250
        .Width = 100
        .Visible = True
        
   End With
   With Picture1
      .AutoRedraw = True
      .AutoSize = True
      .Height = 220
      .Left = 20
      .ScaleMode = 3          'Pixel
      .Top = 20
      .Width = 200
    End With
   With Picture2
      .AutoRedraw = True
      .AutoSize = True
      .Height = 220
      .Left = 200
      .ScaleMode = 3          'Pixel
      .Top = 20
      .Width = 180
   End With
    Picture1.Picture = LoadPicture(App.Path + "\鸟.bmp")
End Sub
