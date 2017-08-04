VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form COMForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "坦克测机校准专用程序__Ver2.0"
   ClientHeight    =   5235
   ClientLeft      =   3720
   ClientTop       =   2715
   ClientWidth     =   10440
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   10440
   Begin VB.CommandButton Command1 
      Caption         =   "刷新串口"
      Height          =   350
      Left            =   9240
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   10380
      TabIndex        =   17
      Top             =   4860
      Width           =   10440
   End
   Begin VB.CommandButton Cmd_reset 
      Caption         =   "复位"
      Height          =   375
      Left            =   9000
      Picture         =   "Form12.frx":58C3A
      TabIndex        =   16
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ComboBox Comb_port 
      Height          =   300
      ItemData        =   "Form12.frx":B1874
      Left            =   840
      List            =   "Form12.frx":B1876
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   660
      Width           =   1095
   End
   Begin VB.ComboBox Comb_bsp 
      Height          =   300
      ItemData        =   "Form12.frx":B1878
      Left            =   840
      List            =   "Form12.frx":B1894
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1035
      Width           =   1095
   End
   Begin VB.ComboBox Comb_databit 
      Height          =   300
      ItemData        =   "Form12.frx":B18CF
      Left            =   840
      List            =   "Form12.frx":B18D6
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1545
      Width           =   1095
   End
   Begin VB.ComboBox Comb_stop 
      Height          =   300
      ItemData        =   "Form12.frx":B18DD
      Left            =   840
      List            =   "Form12.frx":B18E4
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ComboBox Comb_check 
      Height          =   300
      ItemData        =   "Form12.frx":B18EB
      Left            =   840
      List            =   "Form12.frx":B18F2
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_PortSwitch 
      Caption         =   "打开串口"
      Height          =   350
      Left            =   9240
      TabIndex        =   4
      Top             =   240
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_Clear 
      Caption         =   "清除"
      Height          =   350
      Left            =   9240
      TabIndex        =   3
      Top             =   3000
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_Send 
      Caption         =   "发送"
      Height          =   375
      Left            =   9000
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text_send 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4320
      Width           =   8655
   End
   Begin VB.TextBox Text_Receive 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1440
      Width           =   8175
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9720
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Image Image10 
      Height          =   6870
      Left            =   11160
      Picture         =   "Form12.frx":B18FC
      Top             =   4440
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  'Fixed Single
      Height          =   6930
      Left            =   10920
      Picture         =   "Form12.frx":B913E
      Top             =   0
      Width           =   3240
   End
   Begin VB.Image Image8 
      Height          =   6870
      Left            =   9360
      Picture         =   "Form12.frx":BFBF2
      Top             =   5160
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   8775
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   1200
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "当前校准档位："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "mA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   2760
      TabIndex        =   21
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "μA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   495
      Left            =   2760
      TabIndex        =   20
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   240
      Picture         =   "Form12.frx":CE128
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   1920
      Top             =   5880
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   1920
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   615
      Left            =   2040
      TabIndex        =   19
      Top             =   720
      Width           =   7095
   End
   Begin VB.Label Label8 
      Caption         =   "COM数量"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6015
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   9360
      Picture         =   "Form12.frx":CF3FF
      Top             =   3360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   270
      Left            =   9000
      Picture         =   "Form12.frx":CF8F1
      Top             =   3360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   8880
      Top             =   240
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "串口号"
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label2 
      Caption         =   "波特率"
      Height          =   255
      Left            =   105
      TabIndex        =   14
      Top             =   1125
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "数据位"
      Height          =   255
      Left            =   105
      TabIndex        =   13
      Top             =   1620
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "停止位"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2115
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "校验位"
      Height          =   375
      Left            =   105
      TabIndex        =   11
      Top             =   2625
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "发送字符"
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   4920
      Width           =   720
   End
End
Attribute VB_Name = "COMForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public num, comnum As Integer
Dim flagstart As Integer
Dim gpflag As Boolean

Private Sub Command2_Click()
Dim i, j       As Integer
 ' Private Sub Command1_Click()
      j = 0
      
      If (MSComm1.PortOpen = True) Then
      MSComm1.PortOpen = False
      End If
      
      For i = 1 To 16
              MSComm1.CommPort = i
              On Error Resume Next
              MSComm1.PortOpen = True
              If Err.Number = 0 Then
                    j = j + 1
                    Combo1.AddItem "COM" & i
                Else
                    MSComm1.PortOpen = False
              End If
      Next i
     ' Combo1.AddItem "COM" & i
      Label8.Caption = "您计算机上有：   " & j & "个COM"
      comnum = j
End Sub

Private Sub Command1_Click()
Call Form_Load

End Sub

'---------
Private Sub Form_Click()
   Label7.FontSize = 10
  Label7.Caption = "欢迎使用坦克测机校准程序"
  If (num > 10) Then
   Label7.FontSize = 10
  Label7.Caption = "陈峰的QQ号：690293499  :)"
  
  End If
  'Print "坦克测机校准程序"
End Sub
Private Sub Form_dblClick()
   Label7.FontSize = 10
  Label7.Caption = "别闹了，开始校准吧 "
  num = num + 1
  
   
End Sub
Private Sub Form_Load()
'---------------------------------
  gpflag = False
  Label7.Caption = ""
  Label11.Caption = ""
  Label12.Caption = "1，注意，若出现80xx的报错 ，即表明当前电脑串口异常，请通过电脑设备管理器检查确认；                2，校准时务必遵照sop说明事项，勿焦勿躁保持一颗淡定的心！                                         3，用户校准顺序：30uA、30uA、50uA、50mA 和50mA档,其中第一步耗时较长。"
  MsgBox "请先确认本机串口可以正常工作", vbExclamation + vbOKOnly, "提醒"

With frmPort1
Dim j  As Integer
j = 0
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
For i = 1 To 16
    On Error Resume Next
   '当运行发生错误时，控件转到紧接着发生错误的语句之后的语句，并在此继续运行
    MSComm1.CommPort = i
    MSComm1.PortOpen = True
    Select Case Err.Number
       Case 0                       '错误号为0(也就是没出错),
         Comb_port.AddItem "COM" & Trim(i)
         MSComm1.PortOpen = False
         j = j + 1
      ' Case 8002                '错误号为8002,也就是无效端口
      '   Comb_port.AddItem "Com" & Trim(i) & " 无效端口"
      '   MSComm1.PortOpen = False
       'Case Else
      '           MSComm1.PortOpen = False
    End Select
    Err = 0     '将错误号置0. 注:Err.Number可以简写为Err ,2者等效
Next
     Label8.Caption = "您计算机上可用串口数量：   " & j & "个"
     
     comnum = j
     
     If (j = 0) Then
     Label8.FontSize = 18
     'Label7.Font = 微软雅黑
     Label8.Caption = "没有可以用的串口，请确认后再次运行本程序!"
     Cmd_PortSwitch.Visible = False
     Command1.Visible = True
     Else
         Cmd_PortSwitch.Visible = True
         Command1.Visible = False
    End If
    
End With
 '-----------------------------------
 If (j > 1) Then
   Comb_port.ListIndex = 1     'Comb_port初始化，默认选中第2个串口显示出来,(此处主要考虑台式PC有一个用不了的COM1)
 Else
   Comb_port.ListIndex = 0     'Comb_port初始化，默认选中第1个串口显示出来
 End If
 
    MSComm1.Settings = "115200,n,8,1"
    
    Call comb_port_click  ' 设定使用的COM口
   ' MSComm1.CommPort = 1
    MSComm1.InputLen = 0
    
    MSComm1.InBufferSize = 512
    MSComm1.InBufferCount = 0
    MSComm1.OutBufferSize = 512
    MSComm1.OutBufferCount = 0
    MSComm1.RThreshold = 1
    MSComm1.SThreshold = 1
    MSComm1.PortOpen = True
                       
  ' Comb_port.ListIndex = 0     '默认显示第一个串口
    Comb_bsp.ListIndex = 0
    Comb_databit.ListIndex = 0
    Comb_stop.ListIndex = 0
    Comb_check.ListIndex = 0
    Image1.Picture = Image3.Picture
    Cmd_PortSwitch.Caption = "关闭串口"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
If (MSComm1.PortOpen = True) Then
    MSComm1.PortOpen = False
End If
End Sub
Private Sub comb_port_click()
    
   '   MsgBox err.Description, vbExclamation + vbOKOnly, "出错信息"
  '  MsgBox "请先确认本机串口可以正常工作", vbExclamation + vbOKOnly, "友情提醒"
    Dim i As Boolean
    If (MSComm1.PortOpen = True) Then
        MSComm1.PortOpen = False
        i = True
    Else
        i = False
    End If
    
    Select Case Comb_port.Text
    Case "COM1"
        MSComm1.CommPort = 1
    Case "COM2"
        MSComm1.CommPort = 2
    Case "COM3"
        MSComm1.CommPort = 3
    Case "COM4"
        MSComm1.CommPort = 4
    Case "COM5"
        MSComm1.CommPort = 5
    Case "COM6"
        MSComm1.CommPort = 6
    Case "COM7"
        MSComm1.CommPort = 7
    Case "COM8"
        MSComm1.CommPort = 8
    Case "COM9"
        MSComm1.CommPort = 9
     Case "COM10"
        MSComm1.CommPort = 10
    Case "COM11"
        MSComm1.CommPort = 11
    Case "COM12"
        MSComm1.CommPort = 12
    End Select
    
    
    If (i = True) Then
        MSComm1.PortOpen = True
    End If
    
        
End Sub

Private Sub Comb_bsp_click()

Dim i As Boolean

If (MSComm1.PortOpen = True) Then
    MSComm1.PortOpen = False
    i = True
Else
    i = False
End If

    Select Case Comb_bsp.ListIndex
    Case 0
           MSComm1.Settings = "115200,n,8,1"
    Case 1
           MSComm1.Settings = "57600,n,8,1"
    Case 2
           MSComm1.Settings = "56000,n,8,1"
    Case 3
           MSComm1.Settings = "38400,n,8,1"
    Case 4
           MSComm1.Settings = "19200,n,8,1"
    Case 5
           MSComm1.Settings = "14400,n,8,1"
    Case 6
           MSComm1.Settings = "9600,n,8,1"
    Case 7
           MSComm1.Settings = "4800,n,8,1"
    End Select

 
If (i = True) Then
    MSComm1.PortOpen = True
End If



            
End Sub

Private Sub cmd_PortSwitch_Click()
    If (MSComm1.PortOpen = True) Then
        MSComm1.PortOpen = False
        Cmd_PortSwitch.Caption = "打开串口"
        Image1.Picture = Image2.Picture
    ElseIf (MSComm1.PortOpen = False) Then
        MSComm1.PortOpen = True
        Cmd_PortSwitch.Caption = "关闭串口"
        Image1.Picture = Image3.Picture
    End If
End Sub
Private Sub Cmd_Clear_Click()
    Text_Receive.Text = ""
End Sub

Private Sub Check_autosend_Click()

Select Case Check_autosend.Value
Case 1
    Text_time.Locked = True
    Timer1.Interval = Int(Text_time.Text)
    Timer1.Enabled = True
Case 0
    Text_time.Locked = False
    Timer1.Enabled = False
End Select

End Sub

Private Sub Cmd_Send_Click()
If (MSComm1.PortOpen = True) Then MSComm1.Output = Text_send.Text
End Sub

Private Sub Cmd_reset_Click()
    Text_send.Text = ""
End Sub

Private Sub MSComm1_OnComm()
Select Case MSComm1.CommEvent
 Dim cf As String
 Dim cft As Boolean
 'cft = False
 
  Case comEvReceive
    
     If (comnum > 200) Then
         Text_Receive.SelStart = 0
         Text_Receive.Text = ""
         comnum = 0
     End If
    '===========================================
     cf = MSComm1.Input
     Label7.FontSize = 18
     cft = False
     
     If InStr(cf, "Calibration_Start") > 0 Then
         Label7.Caption = "校准开始"
         Label11.Caption = "当前校准档位："
         flagstart = 1
         cft = True
         Image9.Picture = Image10.Picture
     ElseIf InStr(cf, "complete") > 0 Then
         Label7.Caption = "校准结束"
         Label11.Caption = ""
         flagstart = 0
         cft = True
         Image9.Picture = Image8.Picture
         
     ElseIf InStr(cf, "MODE0") > 0 Then
         Image4.Picture = Image6.Picture
         Image5.Picture = Image7.Picture
         Label9.ForeColor = &H80000012
         Label10.ForeColor = &H8000000C '&H80000012&
         Label7.Caption = "IO VDD uA校准,请确认电流为30uA......"
         If (flagstart > 3) Then        '因为uA档位只会校准2次，故flagstart超过2则意味着异常
                 Label7.Caption = "校准流程异常，建议重新开始"
         End If
         flagstart = flagstart + 1
         cft = True
         cf = "--微安档校准--"
      
      ElseIf InStr(cf, "MODE1") > 0 Then
         Image4.Picture = Image6.Picture
         Image5.Picture = Image7.Picture
         Label9.ForeColor = &H80000012
         Label10.ForeColor = &H8000000C '&H80000012&
         Label7.Caption = "VDD uA校准，请确认电流为30uA......"
         cft = True
         cf = "==微安档校准=="
       
       ElseIf InStr(cf, "MODE2") > 0 Then
         Image4.Picture = Image6.Picture
         Image5.Picture = Image7.Picture
         Label9.ForeColor = &H80000012
         Label10.ForeColor = &H8000000C '&H80000012&
         Label7.Caption = "VDD uA校准，请确认电流为50uA......"
         cft = True
         cf = "==微安档校准=="
       
       ElseIf InStr(cf, "MODE3") > 0 Then
         Image4.Picture = Image7.Picture
         Image5.Picture = Image6.Picture
         Label9.ForeColor = &H8000000C
         Label10.ForeColor = &H80000012
         Label7.Caption = "VDD mA校准，请确认电流为50mA......"
         cft = True
         cf = "==毫安档校准=="
        ElseIf InStr(cf, "MODE4") > 0 Then
         Image4.Picture = Image7.Picture
         Image5.Picture = Image6.Picture
         Label9.ForeColor = &H8000000C
         Label10.ForeColor = &H80000012
         Label7.Caption = "IO VDD mA校准，请确认电流为50mA......"
         cft = True
         cf = "==毫安档校准=="
      End If
        
        
      If (cft = True) Then
         Text_Receive.Text = Text_Receive.Text + vbCr + vbLf
         '光标换行
         Text_Receive.Text = Text_Receive.Text + cf
         'Text_Receive.SelStart = Len(Text_Receive)
         comnum = comnum + 1
         '统计接收字符个数，超过特定数量后会清空窗口
         Text_Receive.Text = Text_Receive.Text + vbCr + vbLf
         
         Text_Receive.SelStart = Len(Text_Receive)
        ' Image9.Picture = Image10.Picture
       Else
       
         Text_Receive.Text = Text_Receive.Text + cf
         Text_Receive.SelStart = Len(Text_Receive)
         '设定输入内容的起始位置
         comnum = comnum + 1
       End If
        
         
         
    Case Else
End Select
End Sub

Private Sub Timer1_Timer()
If (MSComm1.PortOpen = True) Then MSComm1.Output = Text_send.Text
End Sub

