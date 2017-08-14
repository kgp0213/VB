Attribute VB_Name = "Hexedit"
'��¼����״̬
Public ComState As String
'ָʾrgb��ɫ�Ƿ�ת
Public rgbChange As Boolean
Public Const CLR_RED As Long = 0
Public Const CLR_GREEN As Long = 1
Public Const CLR_BLUE As Long = 2
Public Const CLR_WHITE As Long = 3

'Public dubuggFlag As Byte   '����˫������
Public m, grayNumflag As Integer 'm���������ͻҽ����鳤�ȣ�grayNumflag����gamma����ʱ�ҽ���ʾ����
Public gstr()

'**********************************
'����ģ��
'**********************************
Public strSendText As String    '�����ı�����
Public bytSendByte() As Byte    '���Ͷ���������

'**********************************
'����ģ��
'**********************************

Public bytReceiveByte() As Byte     '���յ����ֽ�
Public intReceiveLen As Integer     '���յ����ֽ���

'**********************************

'**********************************
'��ʾģ��
'**********************************

Public strHex As String         'ʮ�����Ʊ���
Public strAscii As String        'ASCII��



'**********************************
'�ַ�����ʾ��ʮ����������ת��Ϊ��Ӧ���ֽڴ�
'����ת������ֽ���
'**********************************

Function strHexToByteArray(strText As String, bytByte() As Byte) As Integer
    
    Dim HexData As Integer          'ʮ������(������)�����ֽڶ�Ӧֵ
    Dim hstr As String * 1          '��λ�ַ�
    Dim lstr As String * 1          '��λ�ַ�
    Dim HighHexData As Integer      '��λ��ֵ
    Dim LowHexData As Integer       '��λ��ֵ
    Dim HexDataLen As Integer       '�ֽ���
    Dim StringLen As Integer        '�ַ�������
    Dim Account As Integer          '����
        
    strTestn = ""                   '���ֵ
    HexDataLen = 0
    strHexToByteArray = 0
    
    StringLen = Len(strText)
    Account = StringLen \ 2
    ReDim bytByte(Account)
    
    For n = 1 To StringLen
    
        Do                                              '����ո�
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
        
        If HighHexData = -1 Or LowHexData = -1 Then     '�����Ƿ��ַ��ж�ת��
            HexDataLen = HexDataLen - 1
            
            Exit For
        Else
            
            HexData = HighHexData * 16 + LowHexData
            bytByte(HexDataLen) = HexData
            HexDataLen = HexDataLen + 1
            
            
        End If
                        
    Next n
    
    If HexDataLen > 0 Then                              '�������һ��ѭ���ı����ֵ
        HexDataLen = HexDataLen - 1
        ReDim Preserve bytByte(HexDataLen)
    Else
        ReDim Preserve bytByte(0)
    End If
    
    
    If StringLen = 0 Then                               '����ǿմ�,�򲻻����ѭ����
        strHexToByteArray = 0
    Else
        strHexToByteArray = HexDataLen + 1
    End If
    
    
End Function


'**********************************
'�ַ���ʾ��ʮ��������ת��Ϊ��Ӧ������
'�����򷵻�  -1
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
        test = -1                                       '������Ϣ
    End If
    ConvertHexChr = test
    
End Function


'**********************************
'���봦��
'������յ����ֽ���,��������ȫ�ֱ���
'bytReceiveRyte()
'**********************************

Public Sub InputManage(bytInput() As Byte, intInputLenth As Integer)

    

   
    Dim n As Integer                                '�����������ʼ��
    
    ReDim Preserve bytReceiveByte(intReceiveLen + intInputLenth)

    For n = 1 To intInputLenth Step 1
        bytReceiveByte(intReceiveLen + n - 1) = bytInput(n - 1)
    Next n
    
    intReceiveLen = intReceiveLen + intInputLenth
    
End Sub

'***********************************
'Ϊ���׼���ı�
'������ȫ�ֱ���
'strText
'strHex
'strAddress
'������������
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
    
    
    strAscii = ""            '���ó�ֵ
    strHex = ""
    strAddress = ""
    
    '*****************************************
    '���16�������ASCII����ַ���
    '*****************************************
    
    
    
    For n = 1 To intReceiveLen
        intValue = bytReceiveByte(n - 1)
        
        If intValue < 32 Or intValue > 128 Then         '����Ƿ��ַ�
            strSingleChr = Chr(46)                      '���ڲ�����ʾ��ASCII��,
        Else                                            '��"."��ʾ
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
        
        If (n Mod intHexWidth) = 0 Then                 '���û���
            strAscii = strAscii + Chr$(13) + Chr$(10)
            strHex = strHex + Chr$(13) + Chr$(10)
        Else
            
        End If
    Next n
    
    '******************************************
    
  
End Sub



