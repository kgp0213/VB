VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4020
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Private Declare Function SetLocalTime Lib "kernel32" _
                (lpSystemTime As SYSTEMTIME) _
                As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
    Dim str_DateTime As String
    Dim DateTime As SYSTEMTIME
    GetLocalTime DateTime
    str_DateTime = Str(DateTime.wYear) + "/" + Str(DateTime.wMonth) + "/" + _
                    Str(DateTime.wDay) + " " + Str(DateTime.wHour) + ":" + _
                    Str(DateTime.wMinute) + ":" + Str(DateTime.wSecond) + "." + _
                    Str(DateTime.wMilliseconds) + Chr(13)
    
    DateTime.wYear = DateTime.wYear - 1
    SetLocalTime DateTime
    Sleep (1000)
    GetLocalTime DateTime
    str_DateTime = str_DateTime + Str(DateTime.wYear) + "/" + Str(DateTime.wMonth) + "/" + _
                    Str(DateTime.wDay) + " " + Str(DateTime.wHour) + ":" + _
                    Str(DateTime.wMinute) + ":" + Str(DateTime.wSecond) + "." + _
                    Str(DateTime.wMilliseconds)
    MsgBox str_DateTime
End Sub
