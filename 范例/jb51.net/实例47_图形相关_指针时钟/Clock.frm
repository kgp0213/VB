VERSION 5.00
Begin VB.Form frmClock 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2820
   ClientLeft      =   4335
   ClientTop       =   2835
   ClientWidth     =   2865
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   Icon            =   "Clock.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton NoUse 
      Cancel          =   -1  'True
      Height          =   240
      Left            =   2400
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1200
      Top             =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1200
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   315
      Left            =   2040
      TabIndex        =   10
      Top             =   2400
      Width           =   195
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   2520
      TabIndex        =   9
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   315
      Left            =   2400
      TabIndex        =   8
      Top             =   480
      Width           =   195
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Top             =   120
      Width           =   195
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   315
      Left            =   2565
      TabIndex        =   6
      Top             =   1200
      Width           =   195
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   315
      Left            =   600
      TabIndex        =   5
      Top             =   2400
      Width           =   195
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   2535
      Width           =   195
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   240
      Shape           =   3  'Circle
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private LastMinute As Integer
Private LastHour As Integer
Private Lastx As Integer
Private Lasty As Integer

Private Sub Form_Load()
  Lastx = 999
End Sub

Private Sub Timer1_Timer()
  Const pi = 3.141592653
Dim T
Dim X As Integer
Dim Y As Integer

T = Now
SEC = Second(T)
Min = Minute(T)
HR = Hour(T)
frmClock.Scale (-16, 16)-(16, -16)
If Min <> lastMin Or HR <> LastHour Then
 LastMinute = Min
 LastHour = HR
 frmClock.Cls
 Lastx = 999
 frmClock.DrawWidth = 2
 frmClock.DrawMode = 13
 h = HR + pi / 60
 X = 5 * Sin(h * pi / 6)
 Y = 5 * Cos(h * pi / 6)
 frmClock.Line (0, 0)-(X, Y)
 X = 8 * Sin(Min * pi / 30)
 Y = 8 * Cos(Min * pi / 30)
 frmClock.Line (0, 0)-(X, Y)
 frmClock.DrawWidth = 1
End If
frmClock.DrawMode = 10
RED = RGB(255, 0, 0)

X = 10 * Sin(SEC * pi / 30)
Y = 10 * Cos(SEC * pi / 30)
If Lastx <> 999 Then
  frmClock.Line (0, 0)-(Lastx, Lasty), RED
End If
  frmClock.Line (0, 0)-(X, Y), RED

 Lastx = X
 Lasty = Y
End Sub

Private Sub NoUse_Click()
Unload Me
End Sub

