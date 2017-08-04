VERSION 5.00
Object = "{AA916CEC-D740-4904-ADF7-F65DFDC248ED}#1.0#0"; "Project1.ocx"
Begin VB.Form testform 
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin Project1.labelshape labelshape2 
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      labelshape2     =   "Caption"
   End
End
Attribute VB_Name = "testform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub labelshape1_GotFocus()

End Sub

Private Sub labelshape2_dbclick()
MsgBox "Ë«»÷"
End Sub
