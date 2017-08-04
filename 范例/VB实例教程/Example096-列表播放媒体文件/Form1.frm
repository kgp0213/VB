VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "列表播放"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   4980
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   4980
      TabIndex        =   1
      Top             =   0
      Width           =   4980
      Begin VB.CommandButton Command1 
         Caption         =   "设置播放列表"
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4695
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   600
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
      _cy             =   6165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Err_Handle     '拦截错误
    Dim i As Integer         'i存储空格位置
    Dim z As Integer         'z存储查找的起始位置
    Dim y As Integer
    Dim FileNames() As String
    'FileNames数组存储划分后的文件目录和文件名称
    z = 1
    y = 0
    CommonDialog1.ShowOpen    '显示打开对话框
    CommonDialog1.FileName = CommonDialog1.FileName & Chr(32)
    '在CommonDialog1的FileName属性值后面添加空格
    
    For i = 1 To Len(CommonDialog1.FileName)
        i = InStr(z, CommonDialog1.FileName, Chr(32))
        '设置i为FileName属性值中空格的位置
            If i = 0 Then Exit For
                ReDim Preserve FileNames(y)
                FileNames(y) = Mid(CommonDialog1.FileName, z, i - z)
                '将FileName属性以空格作为划分标志
                '分成若干部分存储到FileNames数组
                z = i + 1
                y = y + 1
    Next

     'Set sPlaylist = WindowsMediaPlayer1.playlistCollection.newPlaylist("vidio")
     Set sPlaylist = WindowsMediaPlayer1.playlistCollection.newPlaylist("audio")
     WindowsMediaPlayer1.currentPlaylist = sPlaylist
     '设置当前播放列表
   For i = 1 To y - 1
      Set sMedia = WindowsMediaPlayer1.mediaCollection.Add(FileNames(0) + FileNames(i))
      '添加选择的文件到播放列表
      WindowsMediaPlayer1.currentPlaylist.insertItem (WindowsMediaPlayer1.currentPlaylist.Count), sMedia
      '添加播放列表项
   Next
   Exit Sub
Err_Handle:
   MsgBox Err.Description
   Exit Sub
End Sub

Private Sub Form_Load()
'初始化各控件
    CommonDialog1.FileName = ""
    'CommonDialog1.Filter = "avi Files|*.avi"
    CommonDialog1.Filter = "Media Files|*.mid"
    CommonDialog1.Flags = cdlOFNAllowMultiselect
    '设置CommonDialog1控件属性
    WindowsMediaPlayer1.settings.autoStart = False
 
End Sub

Private Sub Form_Resize()
'当窗口大小发生变化时自动调整WindowsMediaPlayer1的大小
    WindowsMediaPlayer1.Width = Form1.Width
    WindowsMediaPlayer1.Height = Form1.Height - 800
End Sub

Private Sub Picture1_Resize()
'当Picture1大小发生变化时自动调整其中的Command1的大小
    Command1.Width = Picture1.Width
    Command1.Height = Picture1.Height
End Sub
