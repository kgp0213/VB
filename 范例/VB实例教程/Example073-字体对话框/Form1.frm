VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "字体"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   4755
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CmnDialog1 
      Left            =   2160
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "设置字体"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开文件"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5953
      _Version        =   393217
      ScrollBars      =   1
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
    With CmnDialog1
            .MaxFileSize = 100
            .CancelError = True
            .Filter = "文件类型(*.rtf)|*.rtf"
            .DialogTitle = "请选择一个RTF格式文件"
            .InitDir = "C:\"
            .Flags = cdlOFNFileMustExist Or cdlOFNReadOnly
        End With
    CmnDialog1.ShowOpen
    oldfilename = CmnDialog1.FileName
    RichTextBox1.Text = ""
    If (Len(oldfilename) > 0) Then
        RichTextBox1.LoadFile (oldfilename)
    End If
    Exit Sub
Err_Handle:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub Command2_Click()
    On Error GoTo Err_Handle
    With CmnDialog1
        .DialogTitle = "字体"
        .Max = 30
        .Min = 8
        .Flags = cdlCFScreenFonts
    End With
    CmnDialog1.ShowFont
    With RichTextBox1
        .SelFontName = CmnDialog1.FontName
        .SelFontSize = CmnDialog1.FontSize
        .SelItalic = CmnDialog1.FontItalic
        .SelBold = CmnDialog1.FontBold
        .SelUnderline = CmnDialog1.FontUnderline
        .SelColor = CmnDialog1.Color
    End With
    Exit Sub
Err_Handle:
    MsgBox Err.Description
    Exit Sub
End Sub
