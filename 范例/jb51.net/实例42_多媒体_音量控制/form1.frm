VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "播放一个Wav"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "设置音量"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      Left            =   1080
      Max             =   255
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   1080
      Max             =   255
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "请输入文件路径"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "右声道"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "左声道"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function waveOutGetVolume Lib "winmm.dll" _
        (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare Function waveOutSetVolume Lib "winmm.dll" _
        (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function waveOutGetDevCaps Lib "winmm.dll" _
        Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Long, _
        lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias _
        "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal _
        uFlags As Long) As Long

Const SND_ASYNC = &H1
Const WAVE_MAPPER = -1&
Const MAXPNAMELEN = 32
Const MMSYSERR_NOERROR = 0

Private Type WAVEOUTCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
    dwFormats As Long
    wChannels As Integer
    dwSupport As Long
End Type
Sub SetVolume()
    Dim lVol As Long
    
    lVol = CLng(HScroll2.Value) * &H100 Or HScroll1.Value
    '设置音量
    If waveOutSetVolume(WAVE_MAPPER, lVol) <> MMSYSERR_NOERROR Then
        MsgBox "音量设置出错"
    End If
End Sub

Private Sub Command1_Click()
    SetVolume
End Sub

Private Sub Command2_Click()
    '播放声音文件
    sndPlaySound Text1.Text, SND_ASYNC
End Sub

Private Sub Form_Load()
    Dim lVol As Long
    Dim tWaveCaps As WAVEOUTCAPS
    Text1.Text = "c:\windows\media\logoff.wav"
    waveOutGetVolume WAVE_MAPPER, lVol
    Debug.Print Hex(lVol)
    HScroll1.Value = (lVol And 255)
    HScroll2.Value = ((lVol \ &H10000) And 255)
End Sub

