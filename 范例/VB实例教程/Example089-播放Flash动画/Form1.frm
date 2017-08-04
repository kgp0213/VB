VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form Form1 
   Caption         =   "Flash动画"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "暂停"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "播放"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   855
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      _cx             =   4199199
      _cy             =   4198564
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   "000000"
      SWRemote        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ShockwaveFlash1.Movie = App.Path & "\ball.swf"
    '读取同一目录下的Flash动画文件
    ShockwaveFlash1.Play
    Command1.Enabled = False
End Sub

Private Sub Command2_Click()
    If Command2.Caption = "暂停" Then
        ShockwaveFlash1.Playing = False
        '停止动画
        Command2.Caption = "继续"
    Else
        ShockwaveFlash1.Playing = True
        '继续动画
        Command2.Caption = "暂停"
    End If
    
End Sub
