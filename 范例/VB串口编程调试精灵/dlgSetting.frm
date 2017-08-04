VERSION 5.00
Begin VB.Form dlgSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置参数"
   ClientHeight    =   3300
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   6015
      Begin VB.TextBox txtTime 
         Height          =   495
         Left            =   1920
         TabIndex        =   8
         Text            =   "1000"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtSetting 
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Text            =   "19200,e,8,1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtPort 
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "ms 发送一次"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "发送时间间隔"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "串口设置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "所用串口(1,2,3,4):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "dlgSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_Click()
    
    dlgSetting.Hide
    Unload dlgSetting

End Sub

Private Sub cmdOk_Click()
        
    On Error GoTo SettingError
        
    intPort = Val(dlgSetting.txtPort.Text)
    intTime = Val(dlgSetting.txtTime.Text)
    strSet = dlgSetting.txtSetting.Text
    
    
    
    If Not frmMain.ctrMSComm.PortOpen Then
        
        frmMain.ctrMSComm.CommPort = intPort
        frmMain.ctrMSComm.Settings = strSet
        frmMain.ctrMSComm.PortOpen = True
    End If
    
    If Not blnAutoSendFlag And Not blnReceiveFlag Then
        frmMain.ctrMSComm.PortOpen = False
    End If
    dlgSetting.Hide
    Unload dlgSetting
    
    Exit Sub
    
SettingError:
    intPort = 2
    intTime = 1000
    strSet = "9600,n,8,1"
    dlgSetting.Show
    dlgSetting.txtPort.Text = str(intPort)
    dlgSetting.txtSetting.Text = strSet
    dlgSetting.txtTime.Text = str(intTime)
    
    MsgBox (Error(Err.Number))
    
End Sub


