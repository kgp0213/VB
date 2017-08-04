VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "媒体文件浏览器"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7485
   StartUpPosition =   3  '窗口缺省
   Begin VB.FileListBox File1 
      Height          =   3690
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   1770
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   6015
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   7858
      _cy             =   10610
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    '关联文件列表框
End Sub

Private Sub Drive1_Change()
    On Error GoTo IFerr          '拦截错误
    Dir1.Path = Drive1.Drive    '关联目录列表框
    Exit Sub
IFerr:                           '如果磁盘错误
    MsgBox ("请确认驱动器是否准备好或者磁盘已经不可用!"), _
            vbOKOnly + vbExclamation
    '弹出注意对话框
    Drive1.Drive = Dir1.Path        '忽略驱动器改变

End Sub

Private Sub File1_Click()
    Me.WindowsMediaPlayer1.URL = Me.File1.Path + "\" + Me.File1.FileName
End Sub

Private Sub Form_Load()
    File1.Pattern = "*.AVI;*.MOV;*.DAT;*.MPG;*.WAV,*.MID;*.QT;*.MPEG"
    '指定File1中显示固定格式的文件
    Me.WindowsMediaPlayer1.settings.autoStart = False
    '不自动播放
    Me.WindowsMediaPlayer1.settings.playCount = 1
    '播放次数为1
End Sub

