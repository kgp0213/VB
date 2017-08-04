VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   4335
   Begin RichTextLib.RichTextBox rchMainText 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3836
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.Label lblCurrentWord 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' Return the word the mouse is over.
Public Function RichWordOver(rch As RichTextBox, X As Single, Y As Single) As String
Dim pt As POINTAPI
Dim pos As Integer
Dim start_pos As Integer
Dim end_pos As Integer
Dim ch As String
Dim txt As String
Dim txtlen As Integer

    ' Convert the position to pixels.
    pt.X = X \ Screen.TwipsPerPixelX
    pt.Y = Y \ Screen.TwipsPerPixelY

    ' Get the character number
    pos = SendMessage(rch.hWnd, EM_CHARFROMPOS, 0&, pt)
    If pos <= 0 Then Exit Function

    ' Find the start of the word.
    txt = rch.Text
    For start_pos = pos To 1 Step -1
        ch = Mid$(rch.Text, start_pos, 1)
        ' Allow digits, letters, and underscores.
        If Not ( _
            (ch >= "0" And ch <= "9") Or _
            (ch >= "a" And ch <= "z") Or _
            (ch >= "A" And ch <= "Z") Or _
            ch = "_" _
        ) Then Exit For
    Next start_pos
    start_pos = start_pos + 1

    ' Find the end of the word.
    txtlen = Len(txt)
    For end_pos = pos To txtlen
        ch = Mid$(txt, end_pos, 1)
        ' Allow digits, letters, and underscores.
        If Not ( _
            (ch >= "0" And ch <= "9") Or _
            (ch >= "a" And ch <= "z") Or _
            (ch >= "A" And ch <= "Z") Or _
            ch = "_" _
        ) Then Exit For
    Next end_pos
    end_pos = end_pos - 1

    If start_pos <= end_pos Then _
        RichWordOver = Mid$(txt, start_pos, end_pos - start_pos + 1)
End Function

Private Sub Form_Load()
    rchMainText.Text = "Visual Basic 100 Excellent Examples" & _
        vbCrLf & vbCrLf & "Chapter 2 Exertion of Visual Basic by sorts" & _
        vbCrLf & vbCrLf & "Example 68 Acquire the Word from TextBox"
End Sub

Private Sub rchMainText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim txt As String

    txt = RichWordOver(rchMainText, X, Y)
    If lblCurrentWord.Caption <> txt Then _
        lblCurrentWord.Caption = txt
End Sub

