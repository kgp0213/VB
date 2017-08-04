VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Find file"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   3375
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command_Exit 
      Caption         =   "退出"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command_Find 
      Caption         =   "查找"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ListBox List_Result 
      Height          =   1680
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text_FileName 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "要查找的文件:"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1

Private Declare Function FindFirstFile Lib "kernel32" _
                Alias "FindFirstFileA" _
                (ByVal lpFileName As String, _
                lpFindFileData As WIN32_FIND_DATA) _
                As Long

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type


Private Declare Function FindNextFile Lib "kernel32" _
                Alias "FindNextFileA" _
                (ByVal hFindFile As Long, _
                lpFindFileData As WIN32_FIND_DATA) _
                As Long

Private Declare Function FindClose Lib "kernel32" _
                (ByVal hFindFile As Long) _
                As Long

Private Sub Command_Exit_Click()
    End
End Sub

Private Sub Command_Find_Click()
    Dim fd As WIN32_FIND_DATA
    Me.List_Result.Clear
    'List_Result用来保存查找结果
    Dim hd As Long
    hd = FindFirstFile(Me.Text_FileName, fd)
    '开始查找
    If hd = INVALID_HANDLE_VALUE Then
        Exit Sub
    End If
    Me.List_Result.AddItem fd.cFileName
    While FindNextFile(hd, fd)
        Me.List_Result.AddItem fd.cFileName
    Wend
    FindClose (hd)
    '关闭查找
End Sub

