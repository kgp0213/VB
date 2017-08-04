Attribute VB_Name = "main"

'**********************************
' 基本设置
'**********************************

Public intPort As Integer       '串行口号
Public strSet As String         '协议设置
Public intTime As Integer       '发送时间间隔

'**********************************

'**********************************
'发送与接收标志
'**********************************

Public blnAutoSendFlag As Boolean   '发送标志
Public blnReceiveFlag As Boolean    '接收标志

'**********************************

'**********************************
'发送模块
'**********************************

Public intOutMode As Integer    '发送模式
Public strSendText As String    '发送文本数据
Public bytSendByte() As Byte    '发送二进制数据

'*********************************
'显示标志
'*********************************


Public intHexChk As Integer         '十六进制编码标志
Public intAsciiChk As Integer       'ASCII码标志
Public intAddressChk As Integer     '地址标志
Public intAdd48Chk As Integer       '4/8位地址标志

'**********************************

