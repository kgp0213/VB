VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   2640
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "                            动态效果2"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "            动态效果1"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "动态效果"
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bMoveFlag As Boolean



Private Sub Form_Load()
        bMoveFlag = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bMoveFlag Then
                Label1.BackColor = RGB(0, 255, 0)
                Label3.BackColor = RGB(0, 0, 255)
                bMoveFlag = False
        End If
End Sub

Private Sub Label1_Click()
Timer1.Enabled = False
Label2.Caption = "是不是象Flash的按钮啊？"

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


        If Not bMoveFlag Then
                Label1.BackColor = RGB(255, 0, 0)
                bMoveFlag = True
        End If
End Sub

Private Sub Label3_Click()
Timer1.Enabled = True
Timer1_Timer

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Not bMoveFlag Then
                Label3.BackColor = RGB(255, 0, 0)
                bMoveFlag = True
        End If
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Label2.Caption & "...哈哈"

End Sub
