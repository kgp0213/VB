VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "保存文件"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   4815
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CmnDialog1 
      Left            =   2160
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "保存文件"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "打开文件"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton CmdSaveAs 
      Caption         =   "另存文件"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   3480
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5741
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldfilename As String
Dim newfilename As String
Private Sub CmdOpen_Click()
    On Error GoTo Err_Handle
    With CmnDialog1
            .MaxFileSize = 100
            .CancelError = True
            .Filter = "文件类型(*.rtf)|*.rtf"
            .DialogTitle = "请选择一个RTF格式文件"
            .InitDir = "C:\"
            .Flags = cdlOFNFileMustExist Or cdlOFNReadOnly
        End With
    CmnDialog1.ShowOpen
    oldfilename = CmnDialog1.filename
    RichTextBox1.Text = ""
    If (Len(oldfilename) > 0) Then
        RichTextBox1.LoadFile (oldfilename)
    End If
    CmdSave.Enabled = True
    CmdSaveAs.Enabled = True
    Exit Sub
Err_Handle:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub CmdSave_Click()
    On Error GoTo Err_Handle
    RichTextBox1.SaveFile oldfilename
    Exit Sub
Err_Handle:
   MsgBox Err.Description
   Exit Sub
End Sub

Private Sub CmdSaveAs_Click()
    On Error GoTo Err_Handle
    With CmnDialog1
        .DialogTitle = "另存为"
        .Filter = "RTF格式文件(*.rtf)|*.rtf"
        .DefaultExt = "rtf"
        .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
     End With
     CmnDialog1.ShowSave
     newfilename = CmnDialog1.filename
     If Len(newfilename) Then
            oldfilename = newfilename
            RichTextBox1.SaveFile (oldfilename)
     End If
     Exit Sub
Err_Handle:
   MsgBox Err.Description
   Exit Sub
End Sub
