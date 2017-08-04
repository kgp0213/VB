VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Snow(1000, 2), Amounty As Integer
    Private Sub Form_Load()
    Form1.Show
    DoEvents
    Randomize
    Amounty = 325
    For J = 1 To Amounty
    Snow(J, 0) = Int(Rnd * Form1.Width)
    Snow(J, 1) = Int(Rnd * Form1.Height)
    Snow(J, 2) = 10 + (Rnd * 20)
    Next J
    Do While Not (DoEvents = 0)
    For LS = 1 To 10
    For I = 1 To Amounty
    OldX = Snow(I, 0): OldY = Snow(I, 1)
    Snow(I, 1) = Snow(I, 1) + Snow(I, 2)
    If Snow(I, 1) > Form1.Height Then
     Snow(I, 1) = 0: Snow(I, 2) = 5 + (Rnd * 30)
     Snow(I, 0) = Int(Rnd * Form1.Width)
     OldX = 0: OldY = 0
    End If
    Coloury = 8 * (Snow(I, 2) - 10): Coloury = 60 + Coloury
    PSet (OldX, OldY), QBColor(0)
    PSet (Snow(I, 0), Snow(I, 1)), RGB(Coloury, Coloury, Coloury)
    Next I
    Next LS
    Loop
    End
    End Sub
