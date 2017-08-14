Attribute VB_Name = "Hexedit"
'记录串口状态
Public ComState As String
'指示rgb颜色是否反转
Public rgbChange As Boolean
Public Const CLR_RED As Long = 0
Public Const CLR_GREEN As Long = 1
Public Const CLR_BLUE As Long = 2
Public Const CLR_WHITE As Long = 3

'Public dubuggFlag As Byte   '计数双击次数
Public m, grayNumflag As Integer 'm计数待发送灰阶数组长度，grayNumflag计数gamma画面时灰阶显示次数
Public gstr()

'**********************************
'发送模块
'**********************************
Public strSendText As String    '发送文本数据
Public bytSendByte() As Byte    '发送二进制数据

'**********************************
'接收模块
'**********************************

Public bytReceiveByte() As Byte     '接收到的字节
Public intReceiveLen As Integer     '接收到的字节数

'**********************************

'**********************************
'显示模块
'**********************************

Public strHex As String         '十六进制编码
Public strAscii As String        'ASCII码



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
'输入处理
'处理接收到的字节流,并保存在全局变量
'bytReceiveRyte()
'**********************************

Public Sub InputManage(bytInput() As Byte, intInputLenth As Integer)

    

   
    Dim n As Integer                                '定义变量及初始化
    
    ReDim Preserve bytReceiveByte(intReceiveLen + intInputLenth)

    For n = 1 To intInputLenth Step 1
        bytReceiveByte(intReceiveLen + n - 1) = bytInput(n - 1)
    Next n
    
    intReceiveLen = intReceiveLen + intInputLenth
    
End Sub

'***********************************
'为输出准备文本
'保存在全局变量
'strText
'strHex
'strAddress
'总行数保存在
'intLine
'***********************************

Public Sub GetDisplayText()

    Dim n As Integer
    Dim intValue As Integer
    Dim intHighHex As Integer
    Dim intLowHex As Integer
    Dim strSingleChr As String * 1
    
    Dim intAddress As Integer
    Dim intAddressArray(8) As Integer
    Dim intHighAddress As Integer
    
    intHexWidth = 16
    
    
    strAscii = ""            '设置初值
    strHex = ""
    strAddress = ""
    
    '*****************************************
    '获得16进制码和ASCII码的字符串
    '*****************************************
    
    
    
    For n = 1 To intReceiveLen
        intValue = bytReceiveByte(n - 1)
        
        If intValue < 32 Or intValue > 128 Then         '处理非法字符
            strSingleChr = Chr(46)                      '对于不能显示的ASCII码,
        Else                                            '用"."表示
            strSingleChr = Chr(intValue)
        End If
        
        strAscii = strAscii + strSingleChr
        
        intHighHex = intValue \ 16
        intLowHex = intValue - intHighHex * 16
        
        If intHighHex < 10 Then
            intHighHex = intHighHex + 48
        Else
            intHighHex = intHighHex + 55
        End If
        If intLowHex < 10 Then
            intLowHex = intLowHex + 48
        Else
            intLowHex = intLowHex + 55
        End If
        
        strHex = strHex + " " + Chr$(intHighHex) + Chr$(intLowHex) + " "
        
        If (n Mod intHexWidth) = 0 Then                 '设置换行
            strAscii = strAscii + Chr$(13) + Chr$(10)
            strHex = strHex + Chr$(13) + Chr$(10)
        Else
            
        End If
    Next n
    
    '******************************************
    
  
End Sub



