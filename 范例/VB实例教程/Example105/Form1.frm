VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "回收站"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command_Empty 
      Caption         =   "清空回收站"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" _
                Alias "SHEmptyRecycleBinA" _
                (ByVal hwnd As Long, _
                ByVal pszRootPath As String, _
                ByVal dwFlags As Long) _
                As Long

Private Const SHERB_NOCONFIRMATION = &H1
Private Const SHERB_NOPROGRESSUI = &H2
Private Const SHERB_NOSOUND = &H4

Private Sub Command_Empty_Click()
    Dim result As Long
    result = SHEmptyRecycleBin(0, "", SHERB_NOCONFIRMATION _
                                   Or SHERB_NOPROGRESSUI _
                                   Or SHERB_NOSOUND)
    If result = 0 Then
        Me.Caption = "操作成功！"
    End If
End Sub
