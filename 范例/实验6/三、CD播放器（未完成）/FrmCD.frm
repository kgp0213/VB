VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmCD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CD播放器"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton ComEnd 
      Caption         =   "退出"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton ComEject 
      Caption         =   "弹出"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton ComStop 
      Caption         =   "停止"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton ComPause 
      Caption         =   "暂停"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton ComNext 
      Caption         =   "下一首"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton ComPrev 
      Caption         =   "上一首"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton ComPlay 
      Caption         =   "播放"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "  CD播放器  "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "正在播放总时间："
      Height          =   180
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "正在播放曲目："
      Height          =   180
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "曲目总数："
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   900
   End
End
Attribute VB_Name = "FrmCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  MMControl1.Visible = False
End Sub

Private Sub ComPlay_Click()
  '在未选择文件时，文件名为空字符
  CommonDialog1.FileName = ""
  '设置文件的过滤方式，可显示文件名为.dat的文件
  CommonDialog1.Filter = "(vcd*.dat)│*.dat"
  '初始的文件过滤方式为*.dat
  CommonDialog1.FilterIndex = 2
  '建立打开文件的通用对话框
  CommonDialog1.ShowOpen
  '打开一个文件后关闭前一此被打开的多媒体设备
  MMControl1.Command = "Close"
    '设置多媒体设备类型为MpegVideo
  MMControl1.DeviceType = "MpegVideo"
    '设置打开的文件为通用对话框中选择的文件
  MMControl1.FileName = CommonDialog1.FileName
    '打开文件
  MMControl1.Command = "Open"
  MMControl1.Command = "Play"
  ComPause.Enabled = True
  ComPlay.Enabled = False
  ComStop.Enabled = True
End Sub

Private Sub MMControl1_StatusUpdate()
  Label2.Caption = "曲目总数：" & MMControl1.Tracks
  Label4.Caption = "曲目播放总时间：" & Trim(Str(Int(MMControl1.Length / 60000))) + "分"
  Label3.Caption = "正在播放曲目：" & Str(MMControl1.Track)
End Sub

Private Sub ComPrev_Click()
  MMControl1.Command = "Prev"
End Sub

Private Sub ComNext_Click()
  MMControl1.Command = "Next"
End Sub

Private Sub ComPause_Click()
  ComPlay.Enabled = True
  MMControl1.Command = "Pause"
  ComPause.Enabled = False
End Sub

Private Sub ComStop_Click()
  MMControl1.Command = "Stop"
  ComStop.Enabled = False
  ComPlay.Enabled = True
End Sub

Private Sub ComEject_Click()
  MMControl1.Command = "Stop"
  MMControl1.Command = "Eject"
  ComPlay.Enabled = True
End Sub

Private Sub ComEnd_Click()
  End
End Sub

Private Sub Form_Unload(Cancel As Integer)
  MMControl1.Command = "Stop"
  MMControl1.Command = "Close"
End Sub


