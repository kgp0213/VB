VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Font"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   2640
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   
Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To Screen.FontCount - 1
        Me.List1.AddItem (Screen.Fonts(i))
    Next
End Sub



