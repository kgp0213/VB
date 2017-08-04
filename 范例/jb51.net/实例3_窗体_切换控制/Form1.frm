VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "¿ØÖÆÌ¨"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3720
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "green"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "yellow"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "red"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form2
Form2.Show
End Sub

Private Sub Command2_Click()
Unload Form2
End Sub

Private Sub Command3_Click()
Form2.BackColor = &HFF&
Form2.Show
End Sub

Private Sub Command4_Click()
Form2.BackColor = &HFFFF&
Form2.Show
End Sub

Private Sub Command5_Click()
Form2.BackColor = &HFF00&
Form2.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub
