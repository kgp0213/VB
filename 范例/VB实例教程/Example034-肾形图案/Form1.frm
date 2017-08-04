VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "肾形图案"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   8520
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "绘图"
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   6495
      Left            =   240
      ScaleHeight     =   6435
      ScaleWidth      =   7995
      TabIndex        =   7
      Top             =   720
      Width           =   8055
   End
   Begin VB.TextBox TextN 
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Text            =   "100"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox TextR 
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Text            =   "100"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox TextY 
      Height          =   375
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox TextX 
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "圆的个数："
      Height          =   180
      Left            =   5160
      TabIndex        =   5
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "半径："
      Height          =   180
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "圆心："
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim x, y As Integer
    Dim i, j As Integer
    Dim th As Single
    Dim X0, Y0, R, N, r1 As Integer
    X0 = Val(TextX.Text)
    Y0 = Val(TextY.Text)
    R = Val(TextR.Text)
    N = Val(TextN.Text)
    th = 3.1415926 * 2 / N
    Picture1.Cls
    For i = 0 To N
        x = R * Cos((i - 1) * th)
        y = R * Sin((i - 1) * th)
        r1 = Abs(x)
        Picture1.Circle (x, y), r1
    Next
End Sub
Private Sub Form_Load()
'初始化绘图环境将Picture1定制成一个中心点坐标为(0,0)的坐标系
    Picture1.Scale (-200, 200)-(200, -200)
End Sub

