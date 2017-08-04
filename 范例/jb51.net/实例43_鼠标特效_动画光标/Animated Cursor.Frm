VERSION 5.00
Begin VB.Form AniCursor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animated Cursor"
   ClientHeight    =   1830
   ClientLeft      =   2955
   ClientTop       =   2475
   ClientWidth     =   3990
   Icon            =   "Animated Cursor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1830
   ScaleWidth      =   3990
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3600
      Top             =   960
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   1320
   End
   Begin VB.CommandButton Animate 
      Caption         =   "&Animate Cursor"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   9
      Left            =   2220
      Picture         =   "Animated Cursor.frx":000C
      Top             =   705
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   8
      Left            =   1740
      Picture         =   "Animated Cursor.frx":0316
      Top             =   705
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   7
      Left            =   1260
      Picture         =   "Animated Cursor.frx":0620
      Top             =   705
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   6
      Left            =   780
      Picture         =   "Animated Cursor.frx":092A
      Top             =   705
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   5
      Left            =   270
      Picture         =   "Animated Cursor.frx":0C34
      Top             =   705
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   4
      Left            =   2205
      Picture         =   "Animated Cursor.frx":0F3E
      Top             =   225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   3
      Left            =   1710
      Picture         =   "Animated Cursor.frx":1248
      Top             =   225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   2
      Left            =   1215
      Picture         =   "Animated Cursor.frx":1552
      Top             =   225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   1
      Left            =   720
      Picture         =   "Animated Cursor.frx":185C
      Top             =   225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   0
      Left            =   270
      Picture         =   "Animated Cursor.frx":1B66
      Top             =   225
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "AniCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Index variable to track current cell in animation
'
Private CustomPtr As Integer



Private Sub Animate_Click()

If Not Timer.Enabled Then 'turn on animation if Timer isn't enabled.
   Timer.Enabled = True
   Me.MousePointer = vbCustom
   Animate.Caption = "Stop Animation"
Else
   'turn off animation if Timer is enabled.
   'reset MouseIcon and index to position 0.
   Timer.Enabled = False
   Me.MousePointer = vbDefault
   Animate.Caption = "Animate Cursor"
   CustomPtr = 0
   Me.MouseIcon = imgPtr(CustomPtr)
End If
   
End Sub

Private Sub Timer1_Timer()
   '
   ' Set MouseIcon to current index, and increment
   ' index (looping back if neccessary).
   ' Use cursors designed for B&W if Win16, or
   ' full color cursors in Win32.
   '
   #If Win16 Then
      Me.MouseIcon = imgPtr(CustomPtr + 10)
   #ElseIf Win32 Then
      Me.MouseIcon = imgPtr(CustomPtr)
   #End If
   CustomPtr = (CustomPtr + 1) Mod 10
End Sub

Private Sub Cancel_Click()

Unload Me 'exit the program

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen



End Sub

Private Sub Timer_Timer()

'set MouseIcon to current index, and increment index (looping back if neccessary).
   
Me.MouseIcon = imgPtr(CustomPtr)
CustomPtr = (CustomPtr + 1) Mod 10
   
End Sub


