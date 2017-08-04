VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Icon"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   3030
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command_View 
      Caption         =   "获取图标"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text_Extension 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text_Extension"
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000   ' system icon index
Private Const SHGFI_LARGEICON = &H0         ' large icon
Private Const SHGFI_SMALLICON = &H1         ' small icon
Private Const ILD_TRANSPARENT = &H1         ' display transparent
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or _
             SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or _
             SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Const MAX_PATH = 260

Private Type SHFILEINFO
   hIcon          As Long
   iIcon          As Long
   dwAttributes   As Long
   szDisplayName  As String * MAX_PATH
   szTypeName     As String * 80
End Type

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
                (ByVal himl As Long, _
                ByVal i As Long, _
                ByVal hDCDest As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal flags As Long) As Long

Private Declare Function SHGetFileInfo Lib "shell32.dll" _
                Alias "SHGetFileInfoA" _
                (ByVal pszPath As String, _
                ByVal dwFileAttributes As Long, _
                psfi As SHFILEINFO, _
                ByVal cbSizeFileInfo As Long, _
                ByVal uFlags As Long) As Long
    
Private Declare Function GetTempPath Lib "kernel32" _
                Alias "GetTempPathA" _
                (ByVal nBufferLength As Long, _
                ByVal lpBuffer As String) _
                As Long

Private Declare Function GetTempFileName Lib "kernel32" _
                Alias "GetTempFileNameA" _
                (ByVal lpszPath As String, _
                ByVal lpPrefixString As String, _
                ByVal wUnique As Long, _
                ByVal lpTempFileName As String) _
                As Long

Private Sub Command_View_Click()
    Dim hIcon As Long
    Dim shinfo As SHFILEINFO
    
    Dim sTmpFile As String
    sTmpFile = CreateTempFile("VBT")
    Dim OldName As String
    Dim NewName As String
    OldName = sTmpFile
    NewName = Left(sTmpFile, Len(sTmpFile) - 3) + Me.Text_Extension.Text
    Call FileCopy(OldName, NewName)
    Kill OldName
    sTmpFile = NewName
    
    hIcon = SHGetFileInfo(sTmpFile, 0&, shinfo, Len(shinfo), _
                          BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    Set Picture1.Picture = LoadPicture()
    Picture1.AutoRedraw = True
    Dim result As Integer
    result = ImageList_Draw(hIcon, shinfo.iIcon, Picture1.hDC, _
                            0, 0, ILD_TRANSPARENT)
    Picture1.Picture = Picture1.Image
    Kill sTmpFile
End Sub

Private Sub Form_Load()
    Me.Text_Extension.Text = ""
End Sub

Private Function CreateTempFile(sPrefix As String) As String
    Dim sTmpPath As String * 512
    Dim sTmpName As String * 576
    Dim nRet As Long

    nRet = GetTempPath(512, sTmpPath)
    If (nRet > 0 And nRet < 512) Then
        nRet = GetTempFileName(sTmpPath, sPrefix, 0, sTmpName)
        If nRet <> 0 Then
            CreateTempFile = Left$(sTmpName, _
                    InStr(sTmpName, vbNullChar) - 1)
        End If
    End If
End Function
