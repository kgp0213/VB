VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   2745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton chu 
      Caption         =   "/"
      Height          =   375
      Index           =   15
      Left            =   2160
      TabIndex        =   16
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cheng 
      Caption         =   "×"
      Height          =   375
      Index           =   14
      Left            =   2160
      TabIndex        =   15
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton jian 
      Caption         =   "－"
      Height          =   375
      Index           =   13
      Left            =   2160
      TabIndex        =   14
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton result 
      Caption         =   "="
      Height          =   375
      Index           =   12
      Left            =   1440
      TabIndex        =   13
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      Height          =   375
      Index           =   11
      Left            =   840
      TabIndex        =   12
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton num0 
      Caption         =   "0"
      Height          =   375
      Index           =   10
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton num3 
      Caption         =   "3"
      Height          =   375
      Index           =   9
      Left            =   1440
      TabIndex        =   10
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton num2 
      Caption         =   "2"
      Height          =   375
      Index           =   8
      Left            =   840
      TabIndex        =   9
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton num1 
      Caption         =   "1"
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton num6 
      Caption         =   "6"
      Height          =   375
      Index           =   6
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton num5 
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   840
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton num4 
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton jia 
      Caption         =   "＋"
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton num9 
      Caption         =   "9"
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton num8 
      Caption         =   "8"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton num7 
      Caption         =   "7"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "0"
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double
Dim b As Double
Dim n As Integer

Private Sub cheng_Click(Index As Integer)
b = a
Text1.Text = a
a = 0
n = 3
End Sub

Private Sub chu_Click(Index As Integer)
b = a
Text1.Text = a
a = 0
n = 4
End Sub

Private Sub Command1_Click(Index As Integer)
a = 0
b = 0
Text1.Text = 0
End Sub

Private Sub jia_Click(Index As Integer)
b = a
Text1.Text = a
a = 0
n = 1
End Sub

Private Sub jian_Click(Index As Integer)
b = a
Text1.Text = a
a = 0
n = 2
End Sub

Private Sub num0_Click(Index As Integer)
a = a * 10
Text1.Text = a
End Sub

Private Sub num1_Click(Index As Integer)
a = a * 10 + 1
Text1.Text = a
End Sub

Private Sub num2_Click(Index As Integer)
a = a * 10 + 2
Text1.Text = a
End Sub

Private Sub num3_Click(Index As Integer)
a = a * 10 + 3
Text1.Text = a
End Sub

Private Sub num4_Click(Index As Integer)
a = a * 10 + 4
Text1.Text = a
End Sub

Private Sub num5_Click(Index As Integer)
a = a * 10 + 5
Text1.Text = a
End Sub

Private Sub num6_Click(Index As Integer)
a = a * 10 + 6
Text1.Text = a
End Sub

Private Sub num7_Click(Index As Integer)
a = a * 10 + 7
Text1.Text = a
End Sub

Private Sub num8_Click(Index As Integer)
a = a * 10 + 8
Text1.Text = a
End Sub

Private Sub num9_Click(Index As Integer)
a = a * 10 + 9
Text1.Text = a
End Sub

Private Sub result_Click(Index As Integer)
Select Case n
Case 1
a = a + b
Text1.Text = a
Case 2
a = b - a
Text1.Text = a
Case 3
a = a * b
Text1.Text = a
Case 4
If a = 0 Then
MsgBox "分母不能为0"
Else
a = b / a
Text1.Text = a
End If
End Select
End Sub
