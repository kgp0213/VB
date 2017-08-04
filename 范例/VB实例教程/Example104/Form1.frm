VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "回收站"
   ClientHeight    =   1230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   3255
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command_Delete 
      Caption         =   "删除"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text_Path 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path:"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Boolean
        hNameMappings As Long
        lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias _
        "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40

Private Sub Command_Delete_Click()
    Dim FileOperation As SHFILEOPSTRUCT
    Dim lReturn As Long
    Dim sSendMeToTheBin As String
    sSendMeToTheBin = Me.Text_Path.Text
    sSendMeToTheBin = sSendMeToTheBin + vbNullChar + vbNullChar
    With FileOperation
        .wFunc = FO_DELETE
        .pFrom = sSendMeToTheBin
        .pTo = vbNullChar   ' Not used
        .fFlags = FOF_ALLOWUNDO
    End With
    lReturn = SHFileOperation(FileOperation)
    If lReturn = 0 Then
        MsgBox "删除成功！"
    End If
End Sub
