VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "图像技巧演示"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   3270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "平向插入"
      Height          =   375
      Left            =   1500
      TabIndex        =   3
      Top             =   3360
      Width           =   1035
   End
   Begin VB.PictureBox Picture2 
      Height          =   2715
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   0
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "展开下落"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   420
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   630
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   2460
      Visible         =   0   'False
      Width           =   2235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
     Dim i As Integer

     For i = 0 To 3200
        Picture2.PaintPicture Picture1, 660, i, Picture1.Width, i + 100
     Next i
End Sub

Private Sub Command2_Click()
    Dim i As Integer

     For i = 0 To (Picture2.Width - Picture1.Width) / 2
        Picture2.PaintPicture Picture1, i, (Picture2.Width - Picture1.Width) / 2, Picture1.Width, i + 300
     Next i
     
     For i = 0 To (Picture2.Width - Picture1.Width) / 2
        Picture2.PaintPicture Picture1, i, (Picture2.Width - Picture1.Width) / 2, Picture1.Width, i + 300
     Next i
End Sub

Private Sub Form_Load()

End Sub
