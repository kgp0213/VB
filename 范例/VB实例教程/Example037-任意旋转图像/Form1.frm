VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "旋转图像"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   364
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   532
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "0"
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "旋转"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3735
      Left            =   2880
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2700
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   176
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   151
      TabIndex        =   0
      Top             =   240
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "旋转角度"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5040
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Pi = 3.14

Private Sub Command1_Click()
    Dim x As Integer, y As Integer   '点在Picture1中的坐标
    Dim X1 As Integer, Y1 As Integer '过渡坐标
    Dim X2 As Double, Y2 As Double   '过渡坐标
    Dim X3 As Double, Y3 As Double   '点经过旋转后在Picture2中的坐标
    Dim angle As Double              '旋转的角度
    Dim radian As Double             '旋转的弧度

    angle = Text1.Text
    '获得旋转的角度
    radian = angle * Pi / 180
    '获得旋转的弧度
    Picture2.Cls
    For x = 0 To Picture1.Width
        For y = 0 To Picture1.Height
            X1 = x - Picture1.Width \ 2
            Y1 = y - Picture1.Height \ 2
            '将坐标轴原点移动到Picture1的中心后原来坐标为(x,y)
            '的点的新坐标为(X1,Y1)
            X2 = X1 * Cos(radian) + Y1 * Sin(radian)
            '获得在更改坐标原点后旋转后该点的X坐标
            Y2 = Y1 * Cos(radian) - X1 * Sin(radian)
            '获得在更改坐标原点后旋转后该点的Y坐标
            X3 = X2 + Picture1.Width \ 2
            '获得恢复坐标原点后旋转后的X坐标
            Y3 = Y2 + Picture1.Height \ 2
             '获得恢复坐标原点后旋转后的Y坐标
           If X3 > 0 And X3 < Picture2.Width - 1 And Y3 > 0 _
              And Y3 < Picture2.Height - 1 Then
                Picture2.PSet (X3, Y3), Picture1.Point(x, y)
                '将Picture1中(x,y)点复制到Picture2的(X3,Y3)点
            End If
    Next y
Next x
End Sub

