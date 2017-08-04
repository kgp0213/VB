VERSION 5.00
Begin VB.Form frmtype 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Typing Exercise"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ForeColor       =   &H00FFFFFF&
   Icon            =   "type.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3255
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4200
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4200
      Top             =   1200
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Start"
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "200"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Times"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Scores"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmtype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim score As Integer
Dim speed As Integer

Sub init()
Label1.Caption = Chr(Int(Rnd * 26) + 49)
speed = Int(Rnd * 100 + 100)
Label1.Left = Int(Rnd * Frame1.Width)
Label1.Top = Frame1.Top
End Sub

Sub init1()
Label6.Caption = Chr(Int(Rnd * 26) + 97)
speed = Int(Rnd * 100 + 100)
Label6.Left = Int(Rnd * Frame1.Width)
Label6.Top = Frame1.Top
End Sub

Private Sub Command1_Click()
init
Timer1.Enabled = True
Timer2.Enabled = True
Command1.Visible = False
Label5.Caption = 200
Label4.Caption = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = Label1.Caption Then
 init
 score = score + 1
 Label4.Caption = score
End If
If Chr(KeyAscii) = Label6.Caption Then
 init1
 score = score + 1
 Label4.Caption = score
End If
End Sub

Private Sub Form_Load()
Randomize
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Timer1_Timer()
Label1.Top = Label1.Top + speed
If Label1.Top > Frame1.Height Then
 init
End If
Label6.Top = Label6.Top + speed
If Label6.Top > Frame1.Height Then
 init1
End If
End Sub

Private Sub Timer2_Timer()
Label5.Caption = Val(Label5.Caption) - 1
If Val(Label5.Caption) <= 0 Then
 Timer1.Enabled = False
 Label1.Caption = ""
 Label6.Caption = ""
 Select Case score
  Case Is <= 80
   MsgBox vbCrLf + "Don't give up,try again!"
  Case Is < 200
   MsgBox vbCrLf + "That's right. Come on!"
  Case Is < 350
   MsgBox vbCrLf + "Continue and you will be top gun!"
  Case Is > 350
   MsgBox vbCrLf + "Congraduation! You have been a top gun!"
 End Select
 Command1.Visible = True
 Label4.Caption = 0
 Label5.Caption = 200
 Timer1.Enabled = False
 Timer2.Enabled = False
End If
End Sub

