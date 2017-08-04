VERSION 5.00
Begin VB.Form frmChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "背景自动转换器"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Listfile 
      Height          =   1620
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Textpath 
      Enabled         =   0   'False
      Height          =   270
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Textintval 
      Height          =   270
      Left            =   1920
      TabIndex        =   2
      Text            =   "3"
      Top             =   300
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   2400
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "停止(&T)"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "开始(&S)"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "位图文件存放路径："
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   1620
   End
   Begin VB.Label lblIntval 
      AutoSize        =   -1  'True
      Caption         =   "变换时间间隔(秒)："
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   1530
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function SystemParametersInfo Lib "user32" _
    Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Dim Flag As Boolean
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_UPDATEINIFILE = &H1      'update Win.ini Constant
Const SPIF_SENDWININICHANGE = &H2   'update Win.ini and tell everyone
Private Sub CmdCancel_Click()
    Dim temp As String
    Flag = False
    temp = Textintval.Text & "000"
    Timer1.Interval = Val(temp)

End Sub

Private Sub CmdOK_Click()
    Dim temp As String
    temp = Textpath
    If temp = "" Then Exit Sub
    If Right$(temp, 1) <> "\" Then
        temp = temp + "\"
    End If
    Listfile.Clear
    Listfile.Tag = temp
    temp = temp + "*.bmp"
    temp = Dir$(temp)
    While temp <> ""
        Listfile.AddItem temp
        temp = Dir$
    Wend
    Listfile.AddItem "None"
    Show
    Listfile.ListIndex = 0
    If Listfile.List(0) = "None" Then
        Flag = False
    Else
        Flag = True
    End If
End Sub

Private Sub Form_Load()
    Dim temp As String
    Flag = False
    temp = Textintval.Text & "000"
    Textpath = App.Path
    Timer1.Interval = Val(temp)
    CmdOK_Click
End Sub


Private Sub Timer1_Timer()
    Dim temp As String
    Dim bmpfile As String
    If Flag Then
        temp = Listfile.Tag
        bmpfile = temp + Listfile.List(Listfile.ListIndex)
        SystemParametersInfo SPI_SETDESKWALLPAPER, 0, bmpfile, SPIF_UPDATEINIFILE
        If Listfile.ListIndex = Listfile.ListCount - 1 Then
            Listfile.ListIndex = -1
        End If
        Listfile.ListIndex = Listfile.ListIndex + 1
    End If
End Sub

