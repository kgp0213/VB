VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The GetVersion function returns the operating system in use.
Private Declare Function GetVersion Lib "kernel32" () As Long

Private Sub Form_Load()
'********************************************************************
'* When the project starts, check the operating system used by
'* calling the GetVersion function.
'********************************************************************
    Dim lngVersion As Long
    lngVersion = GetVersion()
    Dim strVersion As String
    strVersion = "Unknown OS"
    If ((lngVersion And &H80000000) = 0) Then
        If ((lngVersion And &HF) = 3) Then
            strVersion = "Windows NT 3.51"
        ElseIf ((lngVersion And &HF) = 4) Then
            strVersion = "Windows NT 4.0"
        ElseIf ((lngVersion And &HF) = 5) Then
            strVersion = "Windows 2000 or Windows XP"
        End If
    ElseIf ((lngVersion And &H80000000) = 1) Then
        If ((lngVersion And &HF) = 3) Then
            strVersion = "Win32s with Windows 3.1"
        ElseIf ((lngVersion And &HF) = 4) Then
            strVersion = "Windows 95, Windows 98, or Windows Me"
        End If
    End If
    MsgBox strVersion
End Sub
