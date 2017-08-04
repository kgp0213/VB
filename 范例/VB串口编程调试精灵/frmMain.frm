VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "串口编程调试精灵"
   ClientHeight    =   7260
   ClientLeft      =   4515
   ClientTop       =   2820
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   8.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmCommWizard"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   9750
   Begin VB.Timer ctrTimer 
      Left            =   4560
      Top             =   120
   End
   Begin VB.Frame Frame3 
      Caption         =   "发送内容"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   2
      Top             =   5880
      Width           =   8175
      Begin VB.ComboBox cboHexAscii 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   6240
         List            =   "frmMain.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtSend 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "接收回显"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   8415
      Begin VB.VScrollBar vsclHexEdit 
         Height          =   2535
         Left            =   5640
         Max             =   5
         TabIndex        =   21
         Top             =   480
         Value           =   2
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.HScrollBar hsclHexEdit 
         Height          =   200
         Left            =   960
         Max             =   5
         TabIndex        =   20
         Top             =   3120
         Value           =   2
         Width           =   3495
      End
      Begin VB.Frame fraHexEditBackground 
         Height          =   2535
         Left            =   600
         TabIndex        =   16
         Top             =   480
         Width           =   4575
         Begin VB.TextBox txtBlank 
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   3240
            TabIndex        =   22
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtHexEditText 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   2280
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtHexEditHex 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   1200
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtHexEditAddress 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   615
         End
      End
      Begin MSComctlLib.Slider sldLenth 
         Height          =   375
         Index           =   0
         Left            =   6600
         TabIndex        =   15
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Min             =   1
         Max             =   32
         SelStart        =   8
         TickFrequency   =   2
         Value           =   8
      End
      Begin VB.CheckBox chkAddress48 
         Caption         =   "4/8位地址"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CheckBox chkAddress 
         Caption         =   "地址"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkAscii 
         Caption         =   "ASCII码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkHex 
         Caption         =   "十六进制码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6720
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtReceive 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3240
         Left            =   200
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   300
         Width           =   6060
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "操作提示"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      Begin VB.CommandButton cmdClear 
         Caption         =   "清除"
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
         Left            =   5640
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdSetting 
         Caption         =   "设置参数"
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
         Left            =   4560
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdReceive 
         Caption         =   "开始接收"
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
         Left            =   3600
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdManualSend 
         Caption         =   "手动发送"
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
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdAutoSend 
         Caption         =   "自动发送"
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
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSCommLib.MSComm ctrMSComm 
      Left            =   6720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   8
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'**********************************
'字符表示的十六进制数转化为相应的整数
'错误则返回  -1
'**********************************

Function ConvertHexChr(str As String) As Integer
    
    Dim test As Integer
    
    test = Asc(str)
    If test >= Asc("0") And test <= Asc("9") Then
        test = test - Asc("0")
    ElseIf test >= Asc("a") And test <= Asc("f") Then
        test = test - Asc("a") + 10
    ElseIf test >= Asc("A") And test <= Asc("F") Then
        test = test - Asc("A") + 10
    Else
        test = -1                                       '出错信息
    End If
    ConvertHexChr = test
    
End Function

'**********************************
'字符串表示的十六进制数据转化为相应的字节串
'返回转化后的字节数
'**********************************

Function strHexToByteArray(strText As String, bytByte() As Byte) As Integer
    
    Dim HexData As Integer          '十六进制(二进制)数据字节对应值
    Dim hstr As String * 1          '高位字符
    Dim lstr As String * 1          '低位字符
    Dim HighHexData As Integer      '高位数值
    Dim LowHexData As Integer       '低位数值
    Dim HexDataLen As Integer       '字节数
    Dim StringLen As Integer        '字符串长度
    Dim Account As Integer          '计数
        
    strTestn = ""                   '设初值
    HexDataLen = 0
    strHexToByteArray = 0
    
    StringLen = Len(strText)
    Account = StringLen \ 2
    ReDim bytByte(Account)
    
    For n = 1 To StringLen
    
        Do                                              '清除空格
            hstr = Mid(strText, n, 1)
            n = n + 1
            If (n - 1) > StringLen Then
                HexDataLen = HexDataLen - 1
                
                Exit For
            End If
        Loop While hstr = " "
        
        Do
            lstr = Mid(strText, n, 1)
            n = n + 1
            If (n - 1) > StringLen Then
                HexDataLen = HexDataLen - 1
                
                Exit For
            End If
        Loop While lstr = " "
        n = n - 1
        If n > StringLen Then
            HexDataLen = HexDataLen - 1
            Exit For
        End If
        
        HighHexData = ConvertHexChr(hstr)
        LowHexData = ConvertHexChr(lstr)
        
        If HighHexData = -1 Or LowHexData = -1 Then     '遇到非法字符中断转化
            HexDataLen = HexDataLen - 1
            
            Exit For
        Else
            
            HexData = HighHexData * 16 + LowHexData
            bytByte(HexDataLen) = HexData
            HexDataLen = HexDataLen + 1
            
            
        End If
                        
    Next n
    
    If HexDataLen > 0 Then                              '修正最后一次循环改变的数值
        HexDataLen = HexDataLen - 1
        ReDim Preserve bytByte(HexDataLen)
    Else
        ReDim Preserve bytByte(0)
    End If
    
    
    If StringLen = 0 Then                               '如果是空串,则不会进入循环体
        strHexToByteArray = 0
    Else
        strHexToByteArray = HexDataLen + 1
    End If
    
    
End Function




'**********************************


Private Sub cboHexAscii_Click()

    
    
    If frmMain.cboHexAscii.Text = "按ASCII码" Then
        intOutMode = 0
        
    Else
        intOutMode = 1
        
    End If
        
End Sub

Private Sub chkAddress_Click()
    
    If chkAddress.Value = 0 Then
        intAddressChk = 0
    Else
        intAddressChk = 1
    End If
    
    Call ScrollRedisplay
    
End Sub

Private Sub chkAddress48_Click()
    
    If chkAddress48.Value = 1 Then
        intAdd48Chk = 1
    Else
        intAdd48Chk = 0
    End If
    
    Call SlideRedisplay
    
End Sub

Private Sub chkAscii_Click()
    
    If chkAscii.Value = 1 Then
        intAsciiChk = 1
    Else
        intAsciiChk = 0
    End If
    
    Call ScrollRedisplay
    
End Sub

Private Sub chkHex_Click()
    
    If chkHex.Value = 0 Then
        intHexChk = 0
    Else
        intHexChk = 1
    End If
    
    Call ScrollRedisplay
    
End Sub

Private Sub cmdAutoSend_Click()
    
    If blnAutoSendFlag Then
        
        frmMain.ctrTimer.Enabled = False
        
        If Not blnReceiveFlag Then
            frmMain.ctrMSComm.PortOpen = False
        End If
        
        frmMain.cmdAutoSend.Caption = "自动发送"
    Else
        If Not frmMain.ctrMSComm.PortOpen Then
            frmMain.ctrMSComm.CommPort = intPort
            frmMain.ctrMSComm.Settings = strSet
            frmMain.ctrMSComm.PortOpen = True
        End If
        
        frmMain.ctrTimer.Interval = intTime
        frmMain.ctrTimer.Enabled = True
        frmMain.cmdAutoSend.Caption = "停止发送"
    End If
        
    
        
    blnAutoSendFlag = Not blnAutoSendFlag
    
    
End Sub

Private Sub cmdClear_Click()

    Dim bytTemp(0) As Byte
    
    ReDim bytReceiveByte(0)
    intReceiveLen = 0
    
    Call InputManage(bytTemp, 0)
    
    Call GetDisplayText
    Call display
    
    
End Sub

Private Sub cmdManualSend_Click()
    
    If Not frmMain.ctrMSComm.PortOpen Then
        frmMain.ctrMSComm.CommPort = intPort
        frmMain.ctrMSComm.Settings = strSet
        frmMain.ctrMSComm.PortOpen = True
    End If
    
        Call ctrTimer_Timer
        
    'If Not blnAutoSendFlag Then
    
    'frmMain.ctrMSComm.PortOpen = False
    'End If
    
End Sub

Private Sub cmdReceive_Click()
    
    
    If blnReceiveFlag Then
        
        If Not blnAutoSendFlag And Not blnReceiveFlag Then
            frmMain.ctrMSComm.PortOpen = False
        End If
        
        frmMain.cmdReceive.Caption = "开始接收"
    Else
        
        If Not frmMain.ctrMSComm.PortOpen Then
            frmMain.ctrMSComm.CommPort = intPort
            frmMain.ctrMSComm.Settings = strSet
            frmMain.ctrMSComm.PortOpen = True
        End If
        
        frmMain.ctrMSComm.InputLen = 0
        frmMain.ctrMSComm.InputMode = 0
        
    
        frmMain.ctrMSComm.InBufferCount = 0
        frmMain.ctrMSComm.RThreshold = 1
        frmMain.cmdReceive.Caption = "停止接收"
    End If
    
    blnReceiveFlag = Not blnReceiveFlag
    
        
End Sub

Private Sub cmdSetting_Click()
    
    dlgSetting.Show
    dlgSetting.txtPort.Text = str(intPort)
    dlgSetting.txtSetting.Text = strSet
    dlgSetting.txtTime.Text = str(intTime)


End Sub

Private Sub ctrMSComm_OnComm()
    
    Dim bytInput() As Byte
    Dim intInputLen As Integer
    
    
    Select Case frmMain.ctrMSComm.CommEvent
        
        
        Case comEvReceive
            If blnReceiveFlag Then
            
                If Not frmMain.ctrMSComm.PortOpen Then
                    frmMain.ctrMSComm.CommPort = intPort
                    frmMain.ctrMSComm.Settings = strSet
                    frmMain.ctrMSComm.PortOpen = True
                End If
                
                '此处添加处理接收的代码
                
                frmMain.ctrMSComm.InputMode = comInputModeBinary
                intInputLen = frmMain.ctrMSComm.InBufferCount
                ReDim bytInput(intInputLen)
                bytInput = frmMain.ctrMSComm.Input
                Call InputManage(bytInput, intInputLen)
                Call GetDisplayText
                Call display
      
                
                If Not blnAutoSendFlag And Not blnReceiveFlag Then
                    frmMain.ctrMSComm.PortOpen = False
                End If
            End If
            
    End Select
    
End Sub

Private Sub ctrTimer_Timer()
    Dim longth As Integer
    
    strSendText = frmMain.txtSend.Text
    If intOutMode = 0 Then
        frmMain.txtReceive.Text = "ascii"
        frmMain.ctrMSComm.Output = strSendText
    Else
        'add code
        longth = strHexToByteArray(strSendText, bytSendByte())
        
        If longth > 0 Then
            frmMain.ctrMSComm.Output = bytSendByte
        End If
        
    End If

End Sub

'*****************************************
'初始化
'*****************************************


Private Sub Form_Load()


    '设置默认发送接收关闭状态
    blnAutoSendFlag = False
    blnReceiveFlag = False
    
    '接收初始化
    intReceiveLen = 0
    
    '默认发送方式为ASCII
    intOutMode = 0
    frmMain.cboHexAscii.Text = "按ASCII码"
    
    '默认显示宽度位数为8
    intHexWidth = 8
    
    frmMain.sldLenth(0).Value = intHexWidth
    
    '默认各复选框处于选定状态
    intHexChk = 1
    intAsciiChk = 1
    intAddressChk = 1
    intAdd48Chk = 1
    
    frmMain.chkAddress.Value = intAddressChk
    frmMain.chkAscii.Value = intAsciiChk
    frmMain.chkHex.Value = intHexChk
    frmMain.chkAddress48.Value = intAdd48Chk
    
    '初始化显示视窗
    frmMain.fraHexEditBackground.Left = frmMain.txtReceive.Left + 30
    frmMain.fraHexEditBackground.Top = frmMain.txtReceive.Top + 30
    frmMain.fraHexEditBackground.Width = frmMain.txtReceive.Width - 60
    frmMain.fraHexEditBackground.Height = frmMain.txtReceive.Height - 60
    
    frmMain.txtHexEditAddress.Top = 0
    frmMain.txtHexEditHex.Top = 0
    frmMain.txtHexEditText.Top = 0
    frmMain.txtBlank.Top = 0
    
    frmMain.txtHexEditAddress.Height = frmMain.fraHexEditBackground.Height
    frmMain.txtHexEditHex.Height = frmMain.fraHexEditBackground.Height
    frmMain.txtHexEditText.Height = frmMain.fraHexEditBackground.Height
    frmMain.txtBlank.Height = frmMain.fraHexEditBackground.Height
    
    '初始化滚动条
    frmMain.vsclHexEdit.Width = 2 * ChrWidth
    frmMain.vsclHexEdit.Top = frmMain.fraHexEditBackground.Top
    frmMain.vsclHexEdit.Left = frmMain.fraHexEditBackground.Left + frmMain.fraHexEditBackground.Width - frmMain.vsclHexEdit.Width
    frmMain.vsclHexEdit.Height = frmMain.fraHexEditBackground.Height
    
    frmMain.hsclHexEdit.Height = ChrHeight
    frmMain.hsclHexEdit.Left = frmMain.fraHexEditBackground.Left
    frmMain.hsclHexEdit.Top = frmMain.fraHexEditBackground.Top + frmMain.fraHexEditBackground.Height - frmMain.hsclHexEdit.Height
    frmMain.hsclHexEdit.Width = frmMain.fraHexEditBackground.Width
    
    
    '设置滚动条最小和最大滚动
    frmMain.vsclHexEdit.Min = 0
    frmMain.vsclHexEdit.SmallChange = 1
    frmMain.vsclHexEdit.LargeChange = 3
    frmMain.vsclHexEdit.Value = 0
    
    frmMain.hsclHexEdit.Min = 0
    frmMain.hsclHexEdit.SmallChange = 1
    frmMain.hsclHexEdit.LargeChange = 3
    frmMain.hsclHexEdit.Value = 0
    
    '显示初始化
  '  Call cmdClear_Click
    
     '初始化串行口
    intPort = 3
    intTime = 1000
    strSet = "19200,e,8,1"
    frmMain.ctrMSComm.InBufferSize = 1024
    frmMain.ctrMSComm.OutBufferSize = 512
    
    
    If Not frmMain.ctrMSComm.PortOpen Then
        frmMain.ctrMSComm.CommPort = intPort
        frmMain.ctrMSComm.Settings = strSet
      '  frmMain.ctrMSComm.PortOpen = True
    End If
    
   ' frmMain.ctrMSComm.PortOpen = False
    
    
End Sub


Private Sub hsclHexEdit_Change()
    intOriginX = -frmMain.hsclHexEdit.Value * ChrWidth
    Call ScrollRedisplay
End Sub

Private Sub sldLenth_Change(Index As Integer)

    intHexWidth = frmMain.sldLenth(0).Value
    Call SlideRedisplay
    

End Sub

Private Sub vsclHexEdit_Change()

    intOriginY = frmMain.vsclHexEdit.Value
    Call ScrollRedisplay
    
End Sub
