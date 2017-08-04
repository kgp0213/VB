VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎使用Visual Basic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   960
      TabIndex        =   0
      Tag             =   "2"
      Top             =   720
      Width           =   2985
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub changcolor(LCnt As Control, color1 As Integer, _
                color2 As Integer, color3 As Integer, _
                color4 As Integer, color5 As Integer, _
                color6 As Integer, color7 As Integer, _
                color8 As Integer)
    Dim tmep As Integer
    tmep = Val(LCnt.Tag)
    Select Case tmep
    Case color1
        LCnt.Tag = color2
    Case color2
        LCnt.Tag = color3
    Case color3
        LCnt.Tag = color4
    Case color4
        LCnt.Tag = color5
    Case color5
        LCnt.Tag = color6
    Case color6
        LCnt.Tag = color7
    Case color7
        LCnt.Tag = color8
    Case color8
        LCnt.Tag = color1
    End Select
    LCnt.ForeColor = QBColor(LCnt.Tag)
    '

End Sub
Private Sub Timer1_Timer()
    changcolor Label1, 2, 3, 4, 5, 6, 7, 8, 9
End Sub
