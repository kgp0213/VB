VERSION 5.00
Begin VB.Form RecycleBin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "资源回收站处理"
   ClientHeight    =   2115
   ClientLeft      =   2715
   ClientTop       =   1770
   ClientWidth     =   4245
   Icon            =   "RecycleBin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "清空资源回收站"
      Height          =   1050
      Left            =   2160
      Picture         =   "RecycleBin.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "检查资源回收站"
      Height          =   1050
      Left            =   600
      Picture         =   "RecycleBin.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "资源回收站中包含 ？ Byte"
      Height          =   195
      Left            =   570
      TabIndex        =   3
      Top             =   1680
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "资源回收站中包含 ？ 个文件"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   1410
      Width           =   2250
   End
End
Attribute VB_Name = "RecycleBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim rbinfo As SHQUERYRBINFO  ' 资源回收站的资讯
    Dim retval As Long           ' 传回值
    ' 初始化 rbinfo 的大小
    rbinfo.cbSize = Len(rbinfo)
    ' 查询资源回收站的内容
    retval = SHQueryRecycleBin("C:\", rbinfo)  ' the path doesn't have to be the root path
    ' 显示资源回收站中目前有多少物件
    If (rbinfo.i64NumItems.LowPart And &H80000000) = &H80000000 Or rbinfo.i64NumItems.HighPart > 0 Then
        Label1 = "C磁盘回收站中有超过 2,147,483,647 个文件"
    Else
        Label1 = "C磁盘回收站中包含 " & rbinfo.i64NumItems.LowPart & " 个文件"
    End If
    ' 显示资源回收站中的物件，占了多少 Bytes。
    If (rbinfo.i64Size.LowPart And &H80000000) = &H80000000 Or rbinfo.i64Size.HighPart > 0 Then
        Label2 = "C磁盘回收站中文件总量超过 2,147,483,647 Byte"
    Else
        Label2 = "C磁盘回收站中文件总量 " & rbinfo.i64Size.LowPart & " Byte"
    End If
End Sub

Private Sub Command2_Click()
    Dim retval As Long  ' return value
    ' 清空所有资源回收站, 不确认
    retval = SHEmptyRecycleBin(RecycleBin.hwnd, "", SHERB_NOCONFIRMATION)
    ' 若有错误讯息出现，则回复资源回收站的图示
    ' 其实这一点不是很需要
    If retval <> 0 Then  ' error
        retval = SHUpdateRecycleBinIcon()
    End If
    Command1_Click
End Sub

