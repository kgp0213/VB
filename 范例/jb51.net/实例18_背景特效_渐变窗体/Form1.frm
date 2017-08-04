VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub colorchange(lphdc As Object, Redval&, Greenval&, Blueval&)
Dim s, r, topv, leftv, rightv, bottomv, colorv As Integer
s = (lphdc.Height / 60)
topv = 0
leftv = 0
rightv = lphdc.Width
bottomv = topv + s
For r = l To 60
lphdc.Line (leftv, topv)-(rightv, bottomv), RGB(Redval, Greenval, Blueval), BF
Redval = Redval - 4
Creenval = Greenval - 4
Blueval = Blueval - 4
If Redval <= 0 Then
Redval = 0
End If
If Greenval <= 0 Then
Greenval = 0
End If
If Blueval <= 0 Then
Blueval = 0
End If
topv = bottomv
bottomv = topv + s
Next
End Sub




Private Sub Form_Resize()
colorchange Form1, 80, 180, 280
End Sub
