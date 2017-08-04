VERSION 5.00
Begin VB.Form frmCopy 
   Caption         =   "Copy Files"
   ClientHeight    =   2220
   ClientLeft      =   1545
   ClientTop       =   1545
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   4890
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Files"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "This Program will create C:\testfolder  then copy all the files into it. The Standard Windows File copy progess bar pops up."
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "frmCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim result As Long, fileop As SHFILEOPSTRUCT
With fileop
        .hwnd = Me.hwnd
        .wFunc = FO_COPY
        .pFrom = "C:\PROGRAM FILES\MICROSOFT VISUAL BASIC\VB.HLP" & vbNullChar & "C:\PROGRAM FILES\MICROSOFT VISUAL BASIC\README.HLP" & vbNullChar & vbNullChar
        .pTo = "C:\testfolder" & vbNullChar & vbNullChar
        .fFlags = FOF_SIMPLEPROGRESS Or FOF_FILESONLY
End With
result = SHFileOperation(fileop)
If result <> 0 Then
        ' Operation failed
        MsgBox Err.LastDllError
Else
        If fileop.fAnyOperationsAborted <> 0 Then
                      MsgBox "Operation Failed"
         End If
End If
End Sub

