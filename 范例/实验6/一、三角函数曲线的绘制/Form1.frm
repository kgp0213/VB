VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "绘制三角函数图形"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "y=Con(x)"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "y=Sin(x)"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "建立坐标系"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Cls
  Form1.Scale (-8, 2)-(8, -2)                     '自定义坐标系
  Line (-7.5, 0)-(7.5, 0)                         '画X轴
  Line (0, 1.7)-(0, -1.7)                         '画Y轴
  CurrentX = 7.6: CurrentY = 0.1: Print "X"       '标识X轴
  CurrentX = 0.5: CurrentY = 1.8: Print "Y"       '标识Y轴
  For i = -7 To 7
    Line (i, 0)-(i, 0.1)                          '在X轴上标记坐标刻度
    CurrentX = i - 0.2: CurrentY = -0.1: Print i  '在X轴上输出数字标识
  Next i
  For i = -1 To 1
    If i <> 0 Then
      CurrentX = -0.7: CurrentY = i + 0.1: Print i '在Y轴上输出数字标识
      Line (0.5, i)-(0, i)                         '在Y轴上标记坐标刻度
    End If
  Next i
End Sub

Private Sub Command2_Click()
  CurrentX = -6.283: CurrentY = 0                '曲线的起点坐标
  For i = -6.283 To 6.283 Step 0.01
  x = i: y = Sin(i)
  Line -(x, y)                                   '绘制正弦曲线
  Next i
  CurrentX = 2.5: CurrentY = 1: Print "y=sin(x)" '输出y=sin(x)
End Sub

Private Sub Command3_Click()
  DrawWidth = 2
  CurrentX = -6.283: CurrentY = 1                 '曲线的起点坐标
  For i = -6.283 To 6.283 Step 0.01
  x = i: y = Cos(i)
  Line -(x, y)                                    '绘制余弦曲线
  Next i
  CurrentX = -7: CurrentY = 1.2: Print "y=cos(x)" '输出y=cos(x)
End Sub

