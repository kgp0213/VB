VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar Form Demo"
   ClientHeight    =   1455
   ClientLeft      =   3630
   ClientTop       =   3420
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1455
   ScaleWidth      =   3015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton cmdCalendar 
      Caption         =   "&Calendar..."
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "&Date:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Let user select date from frmCalendar
Private Sub cmdCalendar_Click()
    Dim UserDate As Date

    UserDate = CVDate(txtDate)
    If frmCalendar.GetDate(UserDate) Then
        txtDate = UserDate
    End If
End Sub

'Exit program
Private Sub cmdExit_Click()
    Unload Me
End Sub

'Initialize date on load
Private Sub Form_Load()
    'Default to today's date
    txtDate = Date
End Sub

