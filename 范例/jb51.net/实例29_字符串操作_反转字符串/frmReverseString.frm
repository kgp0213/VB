VERSION 5.00
Begin VB.Form frmReverseString 
   Caption         =   "Form1"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReverse 
      Caption         =   "&Reverse"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblOutput 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "frmReverseString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReverse_Click()
'
    With frmReverseString
        If .txtInput.Text = "" Then Exit Sub
        .lblOutput.Caption = ReverseString(frmReverseString.txtInput.Text)
    End With
    '
End Sub

Private Sub Form_Load()
'
    With frmReverseString
        .Caption = "Reverse String"
        .txtInput.Text = "always look at the bright side of live"
        .lblOutput.Caption = ""
    End With
    '
End Sub
Private Function ReverseString(strSource As String) As String
'
    Dim pos As Integer          'position
    Dim strDummy As String      'dummy string
    Dim intC As Integer         'counter
    Const Space As String = " " 'space
    '
    pos = Len(strSource)
    strDummy = ""
    For intC = Len(strSource) To 1 Step -1
        If Mid(strSource, intC, 1) = Space Then
            strDummy = strDummy & Mid(strSource, intC, pos - intC + 1)
            pos = intC
        End If
    Next intC
    strDummy = strDummy & Left(strSource, pos - intC)
    ReverseString = Right(strDummy, Len(strDummy) - 1)
    '
End Function
