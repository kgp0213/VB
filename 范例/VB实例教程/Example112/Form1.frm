VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   3495
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim fso As Object
Dim fso_file As Object

Private Sub Form_Load()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fso_file = fso.OpenTextFile(App.Path + "\test.txt", ForReading, True)
    Do While Not fso_file.AtEndOfStream
        Me.Print fso_file.ReadLine
    Loop
    fso_file.Close
End Sub

Private Sub Form_Terminate()
    Set fso_file = fso.OpenTextFile(App.Path + "\test.txt", ForAppending, True)
    fso_file.WriteLine ("new line")
    fso_file.Close
    Set fso_file = Nothing
    Set fso = Nothing
End Sub
