VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   2520
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Windows declarations
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_FINDSTRING = &H14C
Private Const CB_ERR = (-1)

'Declarations for alternate code (see comments below)
'Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const CB_SETCURSEL = &H14E

'Private flag
Private m_bEditFromCode As Boolean


Private Sub Form_Load()
    Dim sSysDir As String, sFile As String

    'Get files from system directory for test list
    Screen.MousePointer = vbHourglass
    sSysDir = Space$(256)
    GetSystemDirectory sSysDir, Len(sSysDir)
    sSysDir = Left$(sSysDir, InStr(sSysDir, Chr$(0)) - 1)
    If Right$(sSysDir, 1) <> "\" Then
        sSysDir = sSysDir & "\"
    End If
    sFile = Dir$(sSysDir & "*.*")
    Do While Len(sFile)
        Combo1.AddItem sFile
        sFile = Dir$
    Loop
    Screen.MousePointer = vbDefault
End Sub

'Certain keystrokes must be handled differently by the Change
'event, so set m_bEditFromCode flag if such a key is detected
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            m_bEditFromCode = True
        Case vbKeyBack
            m_bEditFromCode = True
    End Select
End Sub

Private Sub Combo1_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String

    'Prevent processing as a result of changes from code
    If m_bEditFromCode Then
        m_bEditFromCode = False
        Exit Sub
    End If
    With Combo1
        'Lookup list item matching text so far
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)
        'If match found, append unmatched characters
        If i <> CB_ERR Then
            'Get full text of matching list item
            strTotal = .List(i)
            'Compute number of unmatched characters
            j = Len(strTotal) - Len(strPartial)
            '
            If j <> 0 Then
                'Append unmatched characters to string
                m_bEditFromCode = True
                .SelText = Right$(strTotal, j)
                'Select unmatched characters
                .SelStart = Len(strPartial)
                .SelLength = j
            Else

                '*** Text box string exactly matches list item ***

                'Note: The ListIndex is still -1. If you want to
                'force the ListIndex to the matching item in the
                'list, uncomment the following line. Note that
                'PostMessage is required because Windows sets the
                'ListIndex back to -1 once the Change event returns.
                'Also note that the following line causes Windows to
                'select the entire text, which interferes if the
                'user wants to type additional characters.
'                PostMessage Combo1.hwnd, CB_SETCURSEL, i, 0
            End If
        End If
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

