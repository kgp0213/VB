VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "对话框"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   5010
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CmnDialog1 
      Left            =   1800
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开文件"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5530
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
Private Sub Command1_Click()
    On Error GoTo Err_Handle
    Dim filename As String
    CmnDialog1.ShowOpen
    filename = CmnDialog1.filename
    RichTextBox1.Text = ""
    If (Len(filename) > 0) Then
        RichTextBox1.LoadFile (filename)
    End If
    Exit Sub
Err_Handle:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub Form_Load()
    With CmnDialog1
            .MaxFileSize = 100
            .CancelError = True
            .Filter = "文件类型(*.rtf)|*.rtf"
            .DialogTitle = "请选择一个RTF格式文件"
            .InitDir = "C:\"
            .Flags = cdlOFNFileMustExist Or cdlOFNReadOnly
        End With
End Sub
