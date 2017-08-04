VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "WaitForSingleObject"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   2925
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "运行记事本"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OpenProcess Lib "kernel32" _
                ( _
                ByVal dwDesiredAccess As Long, _
                ByVal bInheritHandle As Long, _
                ByVal dwProcessId As Long _
                ) As Long

Private Declare Function CloseHandle Lib "kernel32" _
                ( _
                ByVal hObject As Long _
                ) As Long
    
Private Declare Function WaitForSingleObject Lib "kernel32" _
                ( _
                ByVal hHandle As Long, _
                ByVal dwMilliseconds As Long _
                ) As Long

Const SYNCHRONIZE = &H100000
Const INFINITE = &HFFFFFFFF

Private Sub Command1_Click()
    Dim pId As Long
    '声明pId变量存储Process Id
    Dim pHnd As Long '
    '声明pHnd变量存储Process Handle
    pId = Shell("Notepad", vbNormalFocus)
    'Shell传回Process Id
    pHnd = OpenProcess(SYNCHRONIZE, 0, pId)
    '取得 Process Handle
    If pHnd <> 0 Then
        Call WaitForSingleObject(pHnd, INFINITE)
        '无限等待，直到程序结束
        Call CloseHandle(pHnd)
    End If
    MsgBox ("记事本已经关闭！")
End Sub
