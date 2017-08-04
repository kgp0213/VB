VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Directory"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   2400
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command_Delete 
      Caption         =   "删除"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command_Create 
      Caption         =   "创建"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateDirectory Lib "kernel32" _
                        Alias "CreateDirectoryA" _
                        (ByVal lpPathName As String, _
                        lpSecurityAttributes As SECURITY_ATTRIBUTES) _
                        As Long

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" _
                        Alias "SHFileOperationA" _
                        (lpFileOp As SHFILEOPSTRUCT) _
                        As Long

Private Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40


Private Sub Command_Create_Click()
    Dim str_Path As String
    Dim sa As SECURITY_ATTRIBUTES
    str_Path = App.Path
    If Right(str_Path, 1) = "\" Then
        Call CreateDirectory(str_Path + "test", sa)
    Else
        Call CreateDirectory(str_Path + "\test", sa)
    End If
End Sub

Private Sub Command_Delete_Click()
    Dim FileOperation As SHFILEOPSTRUCT
    Dim str_Path As String
    Dim result As Long
    str_Path = App.Path
    If Right(str_Path, 1) = "\" Then
        str_Path = str_Path + "test"
    Else
        str_Path = str_Path + "\test"
    End If
    str_Path = str_Path & vbNullChar & vbNullChar
    With FileOperation
        .hwnd = Me.hwnd
        .wFunc = FO_DELETE
        .pFrom = str_Path
        .pTo = vbNullChar
        .fFlags = FOF_ALLOWUNDO
        .lpszProgressTitle = "正在删除文件夹"
    End With
    result = SHFileOperation(FileOperation)
End Sub
