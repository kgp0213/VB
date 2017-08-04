VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "绘制图形"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   4275
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "随机射线"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清屏"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "同心圆"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim X As Integer, Y As Integer, r As Integer, r1 As Integer
  ScaleMode = 6 '坐标度量单位为mm
  DrawWidth = 2 '边线的宽度为2像素
  X = Int(Picture1.ScaleWidth / 2) '绘制的圆不超过窗体
  Y = Int(Picture1.ScaleHeight / 2)
  If ScaleWidth > ScaleHeight Then
    r = Y
  Else
    r = X
  End If
  For r1 = 0 To r
    '绘制同心圆，半径r1逐渐增加
    Picture1.Circle (X, Y), r1, RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
  Next
End Sub

'清除图片框中的内容
Private Sub Command2_Click()
  Picture1.Cls
  Picture1.Scale
End Sub

Private Sub Command3_Click()
  Dim i As Integer
  Picture1.Scale (-320, 240)-(320, -240)
  For i = 1 To 100
  X = 320 * Rnd      '产生X轴
  If Rnd < 0.5 Then X = -X
  Y = 240 * Rnd      '产生Y轴
  If Rnd < 0.5 Then Y = -Y
  colorcode = 15 * Rnd  '产生彩色代码
  Picture1.Line (0, 0)-(X, Y), QBColor(colorcode)
  Next i
End Sub

