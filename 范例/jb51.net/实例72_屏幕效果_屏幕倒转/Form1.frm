VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "非法用户"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4155
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1200
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "终止"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "Form1.frx":030A
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
   Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Const WM_NCLBUTTONDBLCLK = &HA3
    Const WM_NCLBUTTONDOWN = &HA1
    Const HTCAPTION = 2
    Const MF_STRING = &H0&
    Const MF_BYCOMMAND = &H0&
      Const SC_CLOSE = &HF060
    Private hMenu As Long
    Private CloseStr As String
       Private a As Integer
   Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Type Size
        cx As Long
        cy As Long
End Type
Private Declare Function GetViewportExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As Size) As Long
Private Declare Function SetViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SCREENSAVERRUNNING = 97
Private Const SPIF_UPDATEINIFILE = &H1
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_BOTTOM = 1
Private o As Size
Private Sub Command1_Click()
k = GetDC(0)
r = StretchBlt(k, 0, 768, 1024, -768, k, 0, 0, 1024, 768, &HCC0020)
  MsgBox "哈哈哈，被骗了^Y^" & Chr(13) & "报仇：找于一算帐！" & Chr(13) & "ltby1@smth", vbOKOnly + vbInformation, "哈哈哈"
Form1.Hide
End
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Altdown = (Shift And vbAltMask) > 0
If KeyCode = vbKeyF4 Then
If Shift = Altdown Then
sndPlaySound "", &H1
End If
End If
End Sub
        Private Sub Form_Load()
    Label1.Caption = "你是非法用户!!!" & Chr$(13) & "  因此会导致严重错误" & Chr$(13) & "      3秒后显示器的显像管会 爆炸 或 出现严重错误!!!" & Chr$(13) & "    点击‘终止’会停止      @#$%^&*"
    a = 3
Label2.Caption = a
b = App.Path
sndPlaySound "\aa.wav", &H1
    SystemParametersInfo SPI_SCREENSAVERRUNNING, True, &H0, SPIF_UPDATEINIFILE
    SetWindowPos Form1.hwnd, -1, 0, 0, 430, 120, &H40
      End Sub
Private Sub Form_Unload(Cancel As Integer)
    SystemParametersInfo SPI_SCREENSAVERRUNNING, False, &H0, SPIF_UPDATEINIFILE

End Sub

Private Sub Timer1_Timer()
a = a - 1
Label2.Caption = a
  If a = 0 Then
  Timer1.Enabled = False
  Command1_Click
    End
    End If
If a = 12 Then Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Command1.Left = Command1.Left + 350
Command1.Top = Command1.Top + 350
If Command1.Left > Form1.Left + Me.Width + 500 Then Command1.Left = 0
If Command1.Top > Form1.Top + Me.Height + 400 Then Command1.Top = 0
End Sub
