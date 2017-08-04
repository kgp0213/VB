VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "文件阅览器"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   5610
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   1320
   End
   Begin VB.CheckBox Check1 
      Caption         =   "自动滚屏"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "读取文件"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9551
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const SCROLL = &HB6

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
On Error GoTo Err_Handle
    With CommonDialog1
            .MaxFileSize = 100
            .CancelError = True
            .Filter = "文件类型(*.rtf)|*.rtf"
            .DialogTitle = "请选择一个RTF格式文件"
            .InitDir = "C:\"
            .Flags = cdlOFNFileMustExist Or cdlOFNReadOnly
        End With

    Dim filename As String
    CommonDialog1.ShowOpen
    filename = CommonDialog1.filename
    RichTextBox1.Text = ""
    If (Len(filename) > 0) Then
        RichTextBox1.LoadFile (filename)
    End If
    Exit Sub
Err_Handle:
    MsgBox Err.Description
    Exit Sub

End Sub

Private Sub Timer1_Timer()
    SendMessage RichTextBox1.hwnd, SCROLL, 0, 1
End Sub
