Attribute VB_Name = "Hexedit"



'**********************************
'接收模块
'**********************************

Public bytReceiveByte() As Byte     '接收到的字节
Public intReceiveLen As Integer     '接收到的字节数

'**********************************




'**********************************
'显示模块
'**********************************

Public strAddress As String     '地址信息
Public strHex As String         '十六进制编码
Public strAscii As String        'ASCII码

'**********************************



Public intHexWidth As Integer       '显示列数


'**********************************

Public intOriginX As Long       '横向原点(像素)
Public intOriginY As Integer    '纵向原点(行)
Public intLine As Integer       '总行数

'**********************************




'**********************************
'显示常量
'**********************************

Public Const ChrWidth = 105             '单位宽度
Public Const ChrHeight = 2 * ChrWidth   '单位高度
Public Const BorderWidth = 210          '预留边界
Public Const LineMax = 16               '最大显示行数



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
    
    
    '***************************************
    '获得地址字符串
    '***************************************
    
    intLine = intReceiveLen \ intHexWidth
    
    If (intReceiveLen - intHexWidth * intLine) > 0 Then
    intLine = intLine + 1
    End If
    
    For n = 1 To intLine
        intAddress = (n - 1) * intHexWidth
        
        If intAdd48Chk = 1 Then
            intHighAddress = 8
        Else
            intHighAddress = 4
        End If
            intAddressArray(0) = intAddress
        For m = 1 To intHighAddress
            intAddressArray(m) = intAddressArray(m - 1) \ 16
        Next m
        For m = 1 To intHighAddress
            intAddressArray(m - 1) = intAddressArray(m - 1) - intAddressArray(m) * 16
        Next m
        For m = 1 To intHighAddress
        
            If intAddressArray(intHighAddress - m) < 10 Then
                intAddressArray(intHighAddress - m) = intAddressArray(intHighAddress - m) + Asc("0")
                
            Else
                intAddressArray(intHighAddress - m) = intAddressArray(intHighAddress - m) + Asc("A") - 10
                
            End If
            strAddress = strAddress + Chr$(intAddressArray(intHighAddress - m))
        Next m
        
        strAddress = strAddress + Chr$(13) + Chr$(10)       '设置换行
            
    Next n
    
    
    '***************************************
End Sub

'*************************************
'显示输出
'*************************************

Public Sub display()

    
    Dim intViewWidth As Long        '横向宽度(像素)
    Dim intViewLine As Integer      '纵向宽度(行)

    Dim strDisplayAddress As String
    Dim strDisplayHex As String
    Dim strDisplayAscii As String
    
    strDisplayAddress = ""
    strDisplayHex = ""
    strDisplayAscii = ""
    
    Dim intStart As Integer
    Dim intLenth As Integer
    
    
    '***************************************
    '调整显示页面大小,设置滚动位置宽度
    '***************************************
   
    
    If intAdd48Chk = 1 Then
        frmMain.txtHexEditAddress.Width = 8 * ChrWidth + BorderWidth
    Else
        frmMain.txtHexEditAddress.Width = 4 * ChrWidth + BorderWidth
    End If
        
    frmMain.txtHexEditHex.Width = intHexWidth * 4 * ChrWidth + BorderWidth
    frmMain.txtHexEditText.Width = intHexWidth * ChrWidth + BorderWidth
    frmMain.txtBlank.Width = BorderWidth
    
    intViewWidth = frmMain.txtHexEditAddress.Width * intAddressChk + frmMain.txtHexEditHex.Width * intHexChk + frmMain.txtHexEditText.Width * intAsciiChk
    
    If intViewWidth <= frmMain.fraHexEditBackground.Width And intLine < LineMax Then
        frmMain.txtBlank.Width = frmMain.fraHexEditBackground.Width - intViewWidth
        frmMain.hsclHexEdit.Visible = False
        frmMain.vsclHexEdit.Visible = False
        intViewWidth = frmMain.fraHexEditBackground.Width
        intViewLine = intLine
        intOriginX = 0
        intOriginY = 0
        
    ElseIf intViewWidth > frmMain.fraHexEditBackground.Width And intLine < LineMax - 1 Then
        frmMain.hsclHexEdit.Visible = True
        frmMain.vsclHexEdit.Visible = False
        frmMain.hsclHexEdit.Width = frmMain.fraHexEditBackground.Width
        intViewLine = intLine
        intOriginY = 0
        If intOriginX > intViewWidth - frmMain.fraHexEditBackground.Width Then
            intOriginX = intViewWidth - frmMain.fraHexEditBackground.Width
        End If
        
    ElseIf intViewWidth < (frmMain.fraHexEditBackground.Width - frmMain.vsclHexEdit.Width) And intLine >= LineMax Then
        frmMain.vsclHexEdit.Visible = True
        frmMain.hsclHexEdit.Visible = False
        frmMain.txtBlank.Width = frmMain.fraHexEditBackground.Width - intViewWidth
        intViewWidth = frmMain.fraHexEditBackground.Width
        
        intViewLine = LineMax
        
        intOriginX = 0
        If intOriginY > intLine - LineMax Then
            intOriginY = intLine - LineMax
        End If
        
    Else
        frmMain.hsclHexEdit.Visible = True
        frmMain.vsclHexEdit.Visible = True
        frmMain.hsclHexEdit.Width = frmMain.fraHexEditBackground.Width - frmMain.vsclHexEdit.Width
        intViewLine = LineMax - 1
        If intOriginX > intViewWidth - frmMain.fraHexEditBackground.Width Then
            intOriginX = intViewWidth - frmMain.fraHexEditBackground.Width
        End If
        If intOriginY > intLine - LineMax + 1 Then
            intOriginY = intLine - LineMax + 1
        End If
    End If
    
    
    
    frmMain.txtHexEditAddress.Left = intOriginX
    frmMain.txtHexEditHex.Left = intOriginX + frmMain.txtHexEditAddress.Width * intAddressChk
    frmMain.txtHexEditText.Left = intOriginX + frmMain.txtHexEditAddress.Width * intAddressChk + frmMain.txtHexEditHex.Width * intHexChk
    frmMain.txtBlank.Left = intOriginX + frmMain.txtHexEditAddress.Width * intAddressChk + frmMain.txtHexEditHex.Width * intHexChk + frmMain.txtHexEditText.Width * intAsciiChk
    
    intStart = intOriginY * (6 + 4 * intAdd48Chk) + 1
    intLenth = intViewLine * (6 + 4 * intAdd48Chk)
    strDisplayAddress = Mid(strAddress, intStart, intLenth)
    
    intStart = intOriginY * (intHexWidth * 4 + 2) + 1
    intLenth = intViewLine * (intHexWidth * 4 + 2)
    strDisplayHex = Mid(strHex, intStart, intLenth)
    
    intStart = intOriginY * (intHexWidth + 2) + 1
    intLenth = intViewLine * (intHexWidth + 2)
    strDisplayAscii = Mid(strAscii, intStart, intLenth)
    
    
    
    '***************************************
    
    '***************************************
    '设置滚动条
    '***************************************
    
    frmMain.vsclHexEdit.Max = intLine - intViewLine
    frmMain.hsclHexEdit.Max = (intViewWidth - frmMain.fraHexEditBackground.Width) \ ChrWidth + 1
    
    
    '***************************************
    
    
    
    
    
    '***************************************
    '显示输出
    '***************************************
    frmMain.txtHexEditHex.Text = strDisplayHex
    frmMain.txtHexEditText.Text = strDisplayAscii
    frmMain.txtHexEditAddress.Text = strDisplayAddress
    
    '***************************************
    
    
End Sub

'******************************************
'文本无变化的刷新
'******************************************

Public Sub ScrollRedisplay()

    Call display
    
End Sub

'******************************************
'文本发生变化的刷新
'******************************************

Public Sub SlideRedisplay()
    
    Call GetDisplayText
    Call display

End Sub



