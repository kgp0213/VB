VERSION 5.00
Begin VB.Form PingTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "连线检查程式"
   ClientHeight    =   1080
   ClientLeft      =   3750
   ClientTop       =   3180
   ClientWidth     =   2355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "检查连线状态"
      Height          =   525
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   1755
   End
End
Attribute VB_Name = "PingTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Delay(HowLong As Date)
    TempTime = DateAdd("s", HowLong, Now)
    While TempTime > Now
        DoEvents '让 windows 去处理其他事
    Wend
End Sub

Private Sub Command1_Click()
    Dim FileFile As Integer
    Dim TestString As String
    
    '产生一个文字档 Test.txt，写入一个 '0' 字
    TestString = "command.com /c echo 0 > " & "c:\Test.txt"
    Shell (TestString), vbHide
    
    '建立一个 Bat 档，在这个 Bat 档中，我们会设定：
    '随便 Ping 一个在 Internet 上的 Server 两次，将结果写入文字档 Test.txt
    '在这里, 我们以 Ping www.edu.cn 为例
    FileFile = FreeFile
    Open ("c:\Test.bat") For Binary As FileFile
    TestString = "ping -n 2 www.edu.cn > " & "c:\Test.txt"
    Put #FileFile, , TestString
    Close FileFile
    
    '================
    '开始检查是否连线
    '================
    '执行我们建立的 Bat 档 --> Ping
    TestString = "command.com /c " & "c:\Test.bat"
    Shell (TestString), vbHide
    '如果 Ping 成功, 写入文字档 Test.txt 的字串长度至少会大于 200
    '不过由于 Ping 的动作会延迟几秒钟，所以，我们让程式等待 5 秒钟
    Delay 5

    If FileLen("c:\Test.txt") > 201 Then
        Call MsgBox("您的电脑目前已经连线到 Internet！", vbInformation)
    Else
        Call MsgBox("您的电脑目前并未连线到 Internet！.", vbInformation)
    End If
    
    '删除我们在程式中产生的二个档案
    Kill "c:\Test.bat"
    Kill "c:\Test.txt"
End Sub

