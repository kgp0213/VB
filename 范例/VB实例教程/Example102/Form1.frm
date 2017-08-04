VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" _
                (ByVal lpExistingFileName As String, _
                ByVal lpNewFileName As String, _
                ByVal bFailIfExists As Long) _
                As Long

Private Sub Form_Load()
    Dim str_Source As String
    Dim str_Dest As String
    Me.CommonDialog1.DialogTitle = "请选择源文件"
    Me.CommonDialog1.ShowOpen
    str_Source = Me.CommonDialog1.FileName
    If str_Source <> "" Then
        Me.CommonDialog1.DialogTitle = "请输入目标文件"
        Me.CommonDialog1.ShowSave
        str_Dest = Me.CommonDialog1.FileName
        If str_Dest <> "" Then
            FileCopy str_Source, str_Dest
            'CopyFile str_Source, str_Dest, true
            'CopyFile str_Source, str_Dest, False
        End If
    End If
End Sub
