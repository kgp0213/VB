VERSION 5.00
Begin VB.Form WinSeek 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seek"
   ClientHeight    =   3870
   ClientLeft      =   2235
   ClientTop       =   1905
   ClientWidth     =   5970
   ForeColor       =   &H00000080&
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3870
   ScaleWidth      =   5970
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2880
      ScaleHeight     =   2895
      ScaleWidth      =   5775
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ListBox lstFoundFiles 
         BackColor       =   &H00FFFFFF&
         Height          =   2220
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label lblCount 
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblfound 
         Caption         =   "&Files Found:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   5055
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.DriveListBox drvList 
         Height          =   300
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
      Begin VB.DirListBox dirList 
         BackColor       =   &H00FFFFFF&
         Height          =   1560
         Left            =   2040
         TabIndex        =   6
         Top             =   960
         Width           =   2775
      End
      Begin VB.FileListBox filList 
         BackColor       =   &H00FFFFFF&
         Height          =   2070
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtSearchSpec 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Text            =   "*.*"
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label lblCriteria 
         AutoSize        =   -1  'True
         Caption         =   "Search &Criteria:"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   720
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E&xit"
      Height          =   720
      Left            =   1920
      TabIndex        =   1
      Top             =   3000
      Width           =   1200
   End
End
Attribute VB_Name = "WinSeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SearchFlag As Integer    ' Used as flag for cancelling, etc.

Private Sub cmdExit_Click()
  If cmdExit.Caption = "E&xit" Then
    End
  Else                ' If Cancel, just end Search.
    SearchFlag = False
  End If
End Sub

Private Sub cmdSearch_Click()
' Initialize for search, then call DirDiver to perform recursive search.
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
  ' Check what the user did last:
  If cmdSearch.Caption = "&Reset" Then  ' If just a reset,
    ResetSearch                         ' initialize and exit.
    txtSearchSpec.SetFocus
    Exit Sub
  End If

  ' Update dirList.Path if it is different from the currently
  ' selected directory, otherwise perform the search.
  If dirList.Path <> dirList.List(dirList.ListIndex) Then
     dirList.Path = dirList.List(dirList.ListIndex)
     Exit Sub         ' Exit so user can take a look before searching.
  End If

  ' Continue with the search.
  Picture2.Move 0, 0
  Picture1.Visible = False
  Picture2.Visible = True

  cmdExit.Caption = "Cancel"

  filList.Pattern = txtSearchSpec.Text
  FirstPath = dirList.Path
  DirCount = dirList.ListCount

  'Start recursive direcory search.
  NumFiles = 0                       ' Reset global foundfiles indicator.
  result = DirDiver(FirstPath, DirCount, "")
  filList.Path = dirList.Path
  MsgBox "Search OK! " + dirList.Path
  cmdSearch.Caption = "&Reset"
  cmdSearch.SetFocus
  cmdExit.Caption = "E&xit"

End Sub

Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
'  Recursively search directories from NewPath down...
'     NewPath is searched on this recursion.
'     BackUp is origin of this recursion.
'     DirCount is number of subdirectories in this directory.
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Integer
  SearchFlag = True             ' Set flag so user can interrupt.
  DirDiver = False              ' Set to TRUE if there is an error.
  retval = DoEvents()           ' check for events (i.e. user Cancels).
  If SearchFlag = False Then
    DirDiver = True
    Exit Function
  End If
  On Local Error GoTo DirDriverHandler
  DirsToPeek = dirList.ListCount            ' How many directories below this?
  Do While DirsToPeek > 0 And SearchFlag = True
    OldPath = dirList.Path                  ' Save old path for next recursion.
    dirList.Path = NewPath
    If dirList.ListCount > 0 Then
    ' Get to the node bottom.
      dirList.Path = dirList.List(DirsToPeek - 1)
      AbandonSearch = DirDiver((dirList.Path), DirCount%, OldPath)
    End If
    ' Go up 1 level in directories.
    DirsToPeek = DirsToPeek - 1
    If AbandonSearch = True Then Exit Function
  Loop
  ' Call function to enumerate files.
  If filList.ListCount Then
    If Len(dirList.Path) <= 3 Then
        ThePath = dirList.Path         ' If at root level, leave as is...
    Else
        ThePath = dirList.Path + "\"   ' otherwise put "\" before file name.
    End If
    For ind = 0 To filList.ListCount - 1        ' Add conforming files in
        entry = ThePath + filList.List(ind)     ' this directory to listbox.
        lstFoundFiles.AddItem entry
        lblCount.Caption = Str$(Val(lblCount.Caption) + 1)
    Next ind
  End If
  If BackUp <> "" Then         ' If there is a superior
      dirList.Path = BackUp    ' directory, move to it.
  End If
  Exit Function
DirDriverHandler:
  If Err = 7 Then         ' If Out of Memory, assume listbox just got full.
    DirDiver = True       ' Create Msg$ and set return value AbandonSearch.
    MsgBox "You've filled the listbox. Search being abandoned..."
    Exit Function         ' Note that EXIT procedure resets ERR to 0.
  Else                    ' Otherwise display error message and quit.
    MsgBox Error
    End
  End If
End Function

Private Sub DirList_Change()
    ' Update File listbox to sync with Dir listbox.
    filList.Path = dirList.Path
End Sub

Private Sub DirList_LostFocus()
    dirList.Path = dirList.List(dirList.ListIndex)
End Sub

Private Sub DrvList_Change()
    On Error GoTo DriveHandler
    dirList.Path = drvList.Drive
    Exit Sub

DriveHandler:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub

Private Sub Form_Load()
    Picture2.Move 0, 0
    Picture2.Width = WinSeek.ScaleWidth
    Picture2.BackColor = WinSeek.BackColor
    lblCount.BackColor = WinSeek.BackColor
    lblCriteria.BackColor = WinSeek.BackColor
    lblfound.BackColor = WinSeek.BackColor
    Picture1.Move 0, 0
    Picture1.Width = WinSeek.ScaleWidth
    Picture1.BackColor = WinSeek.BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub ResetSearch()
' Reinitialize before starting a new search.
    lstFoundFiles.Clear
    lblCount.Caption = 0
    SearchFlag = False                  ' Flag indicating search in progress.
    Picture2.Visible = False
    cmdSearch.Caption = "&Search"
    cmdExit.Caption = "E&xit"
    Picture1.Visible = True
    dirList.Path = CurDir$: drvList.Drive = dirList.Path ' Reset DOS path.
End Sub

Private Sub txtSearchSpec_Change()
' Update file list box if user changes pattern.
    filList.Pattern = txtSearchSpec.Text
End Sub

Private Sub txtSearchSpec_GotFocus()
    txtSearchSpec.SelStart = 0      ' Highlight the current entry.
    txtSearchSpec.SelLength = Len(txtSearchSpec.Text)
End Sub

