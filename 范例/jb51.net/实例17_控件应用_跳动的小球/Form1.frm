VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2550
   FillColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   2550
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   1800
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   240
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   960
      Shape           =   3  'Circle
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If Shape1.Top >= 500 Then
   Shape1.Top = Shape1.Top - 200
   Else
   Timer1.Enabled = False
   Timer2.Enabled = True
   End If
End Sub

Private Sub Timer2_Timer()
If Shape1.Top <= 1000 Then
Shape1.Top = Shape1.Top + 200
Else
Timer1.Enabled = True
Timer1.Enabled = False
End If
End Sub
