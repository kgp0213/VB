VERSION 5.00
Begin VB.Form UserDlg 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UserDlg"
   ClientHeight    =   3495
   ClientLeft      =   1905
   ClientTop       =   2655
   ClientWidth     =   6855
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3495
   ScaleWidth      =   6855
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "&Exit Setup"
      Height          =   372
      Left            =   3456
      TabIndex        =   4
      Top             =   2976
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   372
      Left            =   1248
      TabIndex        =   3
      Top             =   2976
      Width           =   1572
   End
   Begin VB.TextBox UserText 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   1590
      Width           =   3735
   End
   Begin VB.TextBox UserText 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   1155
      Width           =   3735
   End
   Begin VB.TextBox UserText 
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label outButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "outButton"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5280
      TabIndex        =   8
      Top             =   2976
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label SourcePath 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "SourcePath"
      ForeColor       =   &H80000008&
      Height          =   204
      Left            =   5280
      TabIndex        =   11
      Top             =   2592
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "To quit Setup, choose the Exit button."
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1248
      TabIndex        =   7
      Top             =   2496
      Width           =   3612
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Serial Number:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   840
      TabIndex        =   10
      Top             =   1650
      Width           =   1572
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Company:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   840
      TabIndex        =   9
      Top             =   1215
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   780
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   1248
      TabIndex        =   5
      Top             =   96
      Width           =   5028
   End
End
Attribute VB_Name = "UserDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    FiletoImplant$ = SourcePath.Tag + "SICONVRT.EXE"    '.EXE file to brand
    NumChars% = 30                          'Maximum # of chars per string
    NumStrings% = 3                         'Number of strings to implant
    
    For i = 1 To NumStrings%                'Implant the strings
        ImplantString$ = UserText(i - 1).Text           'User input
        SearchString$ = String$(NumChars%, 87 + i)      'Start with X
        Branded% = Implant(FiletoImplant$, ImplantString$, SearchString$, NumChars%)
        If Branded% <> True Then
            MsgBox "This copy is already registered to another user.", 48, UserDlg.Caption
            UserText(0).SetFocus
            UserText(0).SelStart = 0
            UserText(0).SelLength = Len(UserText(0).Text)
        End If
    Next i

    outButton.Caption = "continue"              'Move on to next step
    'UserDlg.Hide
End Sub

Private Sub Command2_Click()
    outButton.Tag = "exit"
    End
End Sub

Private Function Implant(FiletoImplant As String, ImplantString As String, SearchString As String, NumChars As Integer) As Integer
'Brands .EXE file with user information.
'FiletoImplant - .EXE file to be implanted
'ImplantString - string to be implanted (e.g., user name)
'SearchString  - string in the .EXE file to be replaced by ImplantString
'                (e.g., Const UserName$ = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX")
'NumChars      - number of characters in SearchString
'Function returns TRUE if successful, FALSE if not

    Const BlockSize = 32768                 'size of block read from disk
    Dim FileData As String                  'string to hold block read from disk
    Dim NumBlocks As Integer                'number of complete blocks in .EXE file
    Dim LeftOver As Integer                 'amount left in partial block
    Dim FileLength As Long                  'length of .EXE file
    Dim BlockPosn As Integer                'block number to be checked
    
    Open FiletoImplant For Binary As #1
    FileLength = LOF(1)
    NumBlocks = FileLength \ BlockSize
    LeftOver = FileLength Mod BlockSize
    FileData = String$(BlockSize, 32)
    BlockPosn = 0
    
    For Index = 1 To NumBlocks              'search the .EXE file for special
        Get #1, , FileData                  'string and record location
        Posn& = InStr(FileData, SearchString)
        If Posn& > 0 Then                   'found it!
            BlockPosn = Index
            Seek 1, Posn& + ((BlockPosn - 1) * BlockSize)
            Exit For
        End If
    Next Index
    
    If BlockPosn = 0 Then                   'didn't find it in regular blocks
        FileData = ""                       'so look in leftovers
        FileData = String$(LeftOver, 32)
        Get #1, , FileData
        Posn& = InStr(FileData, SearchString)
        If Posn& = 0 Then                   'string still not found
            Close #1
            Implant = False                 'exit function, return FALSE
            Exit Function
        End If
        Seek 1, Posn&                       'found it in leftovers!
    End If

    temp$ = Space$(NumChars)                'temp space for user info
    LSet temp$ = ImplantString
    Put #1, , temp$                         'brand the .EXE file with user info
    Close #1                                'close file if all strings implanted
    Implant = True                          'end the function
End Function

