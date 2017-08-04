VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HEX 2 BIN Tool V1.0 Beta                          hyc 2011-01-15"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   8775
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame2 
      Caption         =   "Offset 地址:"
      Height          =   735
      Left            =   3360
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
      Begin VB.TextBox OffsetAddr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "H"
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   340
         Width           =   135
      End
   End
   Begin VB.CommandButton Help_Button 
      Caption         =   "帮助?"
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "未使用字节填充为:"
      Height          =   735
      Left            =   1200
      TabIndex        =   11
      Top             =   1560
      Width           =   1935
      Begin VB.OptionButton Option1 
         Caption         =   "FF"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "00"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Convert_Button 
      Caption         =   "转    换"
      Height          =   615
      Left            =   5280
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton BIN_Open 
      Caption         =   "保存为"
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox BIN_PATH 
      Height          =   350
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   6495
   End
   Begin VB.CommandButton HEX_Open 
      Caption         =   "打  开"
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox HEX_PATH 
      Height          =   350
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   6495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "目标BIN文件:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "源HEX文件:"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Dim ReadFileByte() As Byte

Dim HexAddrMin, HexAddrMax As Long



Private Sub HEX_Open_Click()
    
    CommonDialog1.Filter = "HEX文件(*.HEX)|*.hex"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.FileName = ""
    'CommonDialog1.InitDir = App.Path
    CommonDialog1.DialogTitle = "加载需要转换的HEX文件..."
    CommonDialog1.ShowOpen
    
    If Dir(LTrim(CommonDialog1.FileName)) = "" Then
        Y = MsgBox("所选文件不存在!", , "警告")
        CommonDialog1.FileName = HEX_PATH.Text
        Exit Sub
    ElseIf CommonDialog1.FileName <> "" Then
        HEX_PATH.Text = CommonDialog1.FileName  '显示判断参数文件名
    End If
    
End Sub

Private Sub BIN_Open_Click()
    CommonDialog1.Filter = "BIN文件(*.BIN)|*.bin"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.FileName = ""
    'CommonDialog1.InitDir = App.Path
    CommonDialog1.DialogTitle = "保存转换后的BIN文件..."
    CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
        BIN_PATH.Text = CommonDialog1.FileName  '显示判断参数文件名
    End If

End Sub


Private Sub Convert_Button_Click()
    Dim buf()  As Byte
    Dim c(), s$
    Dim n As Long
    If HEX_PATH.Text = "" Then
        MsgBox ("没有选择需要转换的HEX文件!")
        Exit Sub
    End If

    If BIN_PATH.Text = "" Then
        MsgBox ("没有选择保存BIN文件路径!")
        Exit Sub
    End If
    
    Convert_Button.Enabled = False
    HEX_Open.Enabled = False
    BIN_Open.Enabled = False

    OffsetAddr.Text = ""

    If HEX2BIN(HEX_PATH.Text) = True Then

        If Dir(BIN_PATH.Text, vbDirectory) <> "" Then
            Kill BIN_PATH.Text
        End If
        
        ReDim buf(HexAddrMax - HexAddrMin - 1) As Byte
        CopyMemory buf(0), ReadFileByte(0), (HexAddrMax - HexAddrMin)
        ReDim ReadFileByte(0) As Byte
        
      '  Open BIN_PATH.Text For Binary As #1
      '  Put #1, , buf
     '   Close #1
       '---------------------------------------
        ReDim c(HexAddrMax - HexAddrMin - 1)
       For n = 0 To (HexAddrMax - HexAddrMin - 1)
       c(n) = buf(n)
       Next
        s = Join(c)
        s = Join(c, ",")
        
     '   MsgBox ("转换完成!")
        Open BIN_PATH.Text For Output As #1
        Print #1, , ";Total:--"
        Print #1, , str((HexAddrMax - HexAddrMin))
        Print #1, , s
        Close #1
    MsgBox ("转换完成!")
    End If
    
    Convert_Button.Enabled = True
    HEX_Open.Enabled = True
    BIN_Open.Enabled = True

End Sub


Public Function HEX2BIN(FileN As String) As Boolean
  
  Dim LineInBuf As String
  Dim FLen, Ilong, HexStar_Add, Section_addr As Long
  Dim Data_Len, Data_type As Byte


    If Get_HEX_Offset_addr(FileN) = False Then
        HEX2BIN = False
        Exit Function
    End If

  
  On Error Resume Next
  Err = 0
  Open FileN For Input As #1
  
  If Err <> 0 Then
    x = MsgBox("所选文件不存在!请重新选择!", , "警告")
    HEX2BIN = False
    Exit Function
  Else
  
    FLen = FileLen(FileN)
    ReDim ReadFileByte(FLen) As Byte

    If Option1(1).Value = True Then
        For Ilong = 0 To FLen
            ReadFileByte(Ilong) = &HFF
        Next
    End If
  
    Section_addr = 0
  
    Do Until EOF(1)
      Line Input #1, LineInBuf
      Data_type = Mid(LineInBuf, 8, 2)
      Data_Len = Val("&H" + (Mid(LineInBuf, 2, 2)))
      HexStar_Add = Int2hex(Val("&H" + (Mid(LineInBuf, 4, 4))))

      Select Case Data_type
        Case 0
            HexStar_Add = HexStar_Add + Section_addr
            
            If Data_Len <> 0 Then
                For i = 0 To (Data_Len - 1) * 2 Step 2
                    ReadFileByte(HexStar_Add - HexAddrMin) = Val("&H" + (Mid(LineInBuf, i + 10, 2)))
                    HexStar_Add = HexStar_Add + 1
                Next
            End If
        Case 1
            Exit Do
        Case 2
            HexStar_Add = Int2hex(Val("&H" + (Mid(LineInBuf, 10, 4))))
            Section_addr = HexStar_Add * 16
        Case 4
            HexStar_Add = Int2hex(Val("&H" + (Mid(LineInBuf, 10, 4))))
            Section_addr = HexStar_Add * 16 * 16 * 16 * 16
      End Select
    Loop
   
    Close #1

    HEX2BIN = True
  End If
  
End Function

Public Function Get_HEX_Offset_addr(FileN As String) As Boolean
    Dim LineInBuf As String
    Dim HexStar_Add, Section_addr As Long
    Dim Data_type, Data_Len As Byte
    
    
    If FileLen(FileN) = 0 Then
        x = MsgBox("所选HEX文件内容为空!请重新选择!", , "警告")
        Get_HEX_Offset_addr = False
        Exit Function
    End If
    
    On Error Resume Next
    Err = 0
    Open FileN For Input As #1
  
    If Err <> 0 Then
        x = MsgBox("所选文件不存在!请重新选择!", , "警告")
        Get_HEX_Offset_addr = False
        Exit Function
    Else
        Section_addr = 0
        HexAddrMax = 0
        HexAddrMin = &H7FFFFFFF
        
        Do Until EOF(1)
            Line Input #1, LineInBuf
            
            Data_type = Mid(LineInBuf, 8, 2)
            Data_Len = Val("&H" + (Mid(LineInBuf, 2, 2)))
            HexStar_Add = Int2hex(Val("&H" + (Mid(LineInBuf, 4, 4))))

            Select Case Data_type
                Case 0
                    HexStar_Add = HexStar_Add + Section_addr
                    If HexStar_Add < HexAddrMin Then
                        HexAddrMin = HexStar_Add                    '取得最小的地址
                    End If
                    If (HexStar_Add + Data_Len) > HexAddrMax Then
                        HexAddrMax = (HexStar_Add + Data_Len)       '取得最大地址
                    End If
                Case 1
                    Exit Do
                Case 2
                    HexStar_Add = Int2hex(Val("&H" + (Mid(LineInBuf, 10, 4))))
                    Section_addr = HexStar_Add * 16
                Case 4
                    HexStar_Add = Int2hex(Val("&H" + (Mid(LineInBuf, 10, 4))))
                    Section_addr = HexStar_Add * 16 * 16 * 16 * 16
            End Select
        Loop
   
        Close #1
        OffsetAddr.Text = Hex(HexAddrMin)
        Get_HEX_Offset_addr = True
    
    End If

End Function

Public Function Int2hex(inte As Long) As Long
    If inte < 0 Then '最高位为1是负数
        Int2hex = 65536 + inte
    Else
        Int2hex = inte
    End If
     
End Function

