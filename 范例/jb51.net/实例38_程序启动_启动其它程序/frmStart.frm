VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Start File"
   ClientHeight    =   2100
   ClientLeft      =   3705
   ClientTop       =   3285
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   4305
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "c:\example.swf"
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start File (Based on Extension)"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "File to Start"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Dim r As Long
 r = StartDoc(Text1.Text)
End Sub

