VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "渐变背景"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "渐变类型"
      Height          =   1575
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
      Begin VB.OptionButton Option3 
         Caption         =   "圆形"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "垂直"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "水平"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub VDither(startForm As Form)
 '垂直渐变
    Dim intLoop As Integer
    startForm.DrawStyle = vbInsideSolid
    startForm.DrawMode = vbCopyPen
    startForm.ScaleMode = vbPixels
    startForm.DrawWidth = 2
    startForm.AutoRedraw = True
    For intLoop = 0 To startForm.ScaleHeight
       startForm.Line (0, intLoop)-(startForm.ScaleWidth, intLoop), RGB(Abs(255 - intLoop), Abs(255 - intLoop), Abs(255 - intLoop)), B
    Next intLoop
End Sub
Sub HDither(startForm As Form)
'水平渐变
    Dim intLoop As Integer
    startForm.DrawStyle = vbInsideSolid
    startForm.DrawMode = vbCopyPen
    startForm.ScaleMode = vbPixels
    startForm.DrawWidth = 2
    startForm.AutoRedraw = True
    For intLoop = 0 To startForm.Width
     startForm.Line (intLoop, 1)-(intLoop, startForm.Height), RGB(Abs(255 - intLoop), 255, 255), B
     '从白色渐变到蓝色
   Next intLoop
End Sub
Sub CDither(startForm As Form)
'圆形渐变
    Dim intLoop As Integer
    startForm.DrawStyle = vbInsideSolid
    '设置线型
    startForm.DrawMode = vbCopyPen
    startForm.ScaleMode = vbPixels
    startForm.DrawWidth = 2
    '设置线宽
    startForm.AutoRedraw = True
    For intLoop = 0 To startForm.Width
    startForm.Circle (startForm.ScaleWidth / 2, startForm.ScaleHeight / 2), intLoop, RGB(Abs(255 - intLoop), Abs(255 - intLoop), Abs(255 - intLoop))
   '从白色渐变到黑色
   Next intLoop
End Sub

Private Sub Form_Load()
    HDither Form1
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        HDither Form1
    End If
End Sub
Private Sub Option2_Click()
    If Option2.Value = True Then
        VDither Form1
    End If
End Sub
Private Sub Option3_Click()
    If Option3.Value = True Then
        CDither Form1
    End If
End Sub

