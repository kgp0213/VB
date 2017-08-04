VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   ScaleHeight     =   1245
   ScaleWidth      =   2400
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Num2Date支持的掩码有:
'   MMDDYY MMDDYYYY
'   DDMMYY DDMMYYYY
'   YYMMDD YYYYMMDD
Function Num2Date(ByVal N As Long, ByVal Fmt As String) As Variant
    Select Case Fmt
        Case "MMDDYY"             '052793
            Num2Date = CDate(N \ 10000 & "/" & N \ 100 Mod 100 & _
                             "/" & N Mod 100)
        Case "MMDDYYYY"           '05271993
            Num2Date = CDate(N \ 1000000 & "/" & N \ 10000 Mod 100 & _
                             "/" & N Mod 10000)
        Case "DDMMYY"             '270593
            Num2Date = CDate(N \ 100 Mod 100 & "/" & N \ 10000 & _
                             "/" & N Mod 100)
        Case "DDMMYYYY"           '27051993
            Num2Date = CDate(N \ 10000 Mod 100 & "/" & N \ 1000000 & _
                             "/" & N Mod 10000)
        Case "YYMMDD", "YYYYMMDD" '930527   19930527
            Num2Date = CDate(N \ 100 Mod 100 & "/" & N Mod 100 & "/" & _
                             N \ 10000)
        Case Else
            Num2Date = Null
    End Select
End Function

'   String2Date支持的掩码有:
'   MMDDYY    MMDDYYYY   MM/DD/YY   MM/DD/YYYY   M/D/Y   M/D/YY   M/D/YYYY
'   DDMMYY    DDMMYYYY   DD/MM/YY   DD/MM/YYYY   DD-MMM-YY   DD-MMM-YYYY
'   YYMMDD    YYYYMMDD   YY/MM/DD   YYYY/MM/DD
Function String2Date(ByVal S As String, _
                     ByVal Fmt As String) As Variant
    Select Case Fmt
        Case "MMDDYY", "MMDDYYYY"      '052793   05271993
            String2Date = CDate(Left(S, 2) & "/" & Mid(S, 3, 2) & "/" & _
                                Mid(S, 5))
        Case "DDMMYY", "DDMMYYYY"      '270593   27051993
            String2Date = CDate(Mid(S, 3, 2) & "/" & Left(S, 2) & "/" & _
                                Mid(S, 5))
        Case "YYMMDD"                  '930527
            String2Date = CDate(Mid(S, 3, 2) & "/" & Right(S, 2) & "/" & _
                                Left(S, 2))
        Case "YYYYMMDD"                '19930527
            String2Date = CDate(Mid(S, 5, 2) & "/" & Right(S, 2) & "/" & _
                                Left(S, 4))
        Case "MM/DD/YY", "MM/DD/YYYY", "M/D/Y", "M/D/YY", "M/D/YYYY", _
               "DD-MMM-YY", "DD-MMM-YYYY"
            String2Date = CDate(S)
        Case "DD/MM/YY", "DD/MM/YYYY"  '27/05/93   27/05/1993
            String2Date = CDate(Mid(S, 4, 3) & Left(S, 3) & Mid(S, 7))
        Case "YY/MM/DD"                '93/05/27
            String2Date = CDate(Mid(S, 4, 3) & Right(S, 2) & _
                                "/" & Left(S, 2))
        Case "YYYY/MM/DD"              '1993/05/27
            String2Date = CDate(Mid(S, 6, 3) & Right(S, 2) & _
                                "/" & Left(S, 4))
        Case Else
            String2Date = Null
    End Select
End Function

Private Sub Form_Load()
    Dim str_Date As String
    str_Date = Str(Num2Date(19980203, "YYYYMMDD")) + Chr(13)
    str_Date = str_Date + Str(String2Date("020398", "MMDDYY"))
    MsgBox str_Date
End Sub
