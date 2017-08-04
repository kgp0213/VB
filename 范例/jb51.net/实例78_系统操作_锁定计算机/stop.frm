VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7620
   ClientLeft      =   1395
   ClientTop       =   1350
   ClientWidth     =   11775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "stop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7620
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Close Programm"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   5
      Top             =   4560
      Width           =   2430
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3045
      Top             =   1155
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      IMEMode         =   3  'DISABLE
      Left            =   5670
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3675
      Width           =   2220
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2625
      Top             =   1155
   End
   Begin VB.Label Label5 
      Caption         =   "Password is : visualbasic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   8400
      TabIndex        =   6
      Top             =   6300
      Width           =   2745
   End
   Begin VB.Label Label4 
      Caption         =   "Push ENTER to type a password."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      TabIndex        =   4
      Top             =   1800
      Width           =   5055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "System Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   6525
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1080
      TabIndex        =   2
      Top             =   2310
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2940
      TabIndex        =   0
      Top             =   3675
      Width           =   2220
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim ter As Boolean

Private Sub Command1_Click()
    ter = False
    DisableCtrlAltDelete (False)
    Unload Form1
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Timer1.Enabled = False
        Timer2.Enabled = True
        Text1.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    DisableCtrlAltDelete (True)
    ter = True
End Sub


Private Sub Form_LostFocus()
    Form1.SetFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If ter = True Then
        Cancel = 5
    Else
    End If
End Sub


Private Sub Timer1_Timer()
    Form1.SetFocus
    Label2.Caption = 5
End Sub


Private Sub Timer2_Timer()
    Label2.Caption = Val(Label2.Caption) - 1 & "  Seconds time to write a password."
    Text1.SetFocus
    If Text1.Text = "visualbasic" Then
        ter = False
        Timer2.Enabled = False
        Timer1.Enabled = False
        Label2.Caption = "5"
        MsgBox "Password accepted, you can close this programm now."
        Text1.Enabled = False
        Command1.Enabled = True
    End If
    If Val(Label2.Caption) = 0 Then
        Timer1.Enabled = True
        Timer2.Enabled = False
        Text1.Text = ""
        Text1.Enabled = False
        Command1.Enabled = False
    End If
End Sub


