VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Recycle"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   4500
   StartUpPosition =   3  '窗口缺省
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox Text_Space 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Text            =   "Text_Space"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox Text_Count 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Text            =   "Text_Count"
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label_Space 
      AutoSize        =   -1  'True
      Caption         =   "占用空间:"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   810
   End
   Begin VB.Label Label_Count 
      AutoSize        =   -1  'True
      Caption         =   "对象数目:"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type ULARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Type SHQUERYRBINFO
    cbSize As Long
    i64Size As ULARGE_INTEGER
    i64NumItems As ULARGE_INTEGER
End Type

Private Declare Function SHQueryRecycleBin Lib "shell32.dll" _
                Alias "SHQueryRecycleBinA" _
                (ByVal pszRootPath As String, _
                pSHQueryRBInfo As SHQUERYRBINFO) _
                As Long

Private Sub Drive1_Change()
    Dim rbinfo As SHQUERYRBINFO '回收站的信息
    Dim result As Long ' 返回值
    
    rbinfo.cbSize = Len(rbinfo)
    '初始化 rbinfo 的大小
    Dim drive As String
    drive = Left(Me.Drive1.drive, 2) + "\"
    result = SHQueryRecycleBin(drive, rbinfo)
    '查询回收站的内容
    
    If (rbinfo.i64NumItems.LowPart And &H80000000) = &H80000000 Or _
        rbinfo.i64NumItems.HighPart > 0 Then
        Me.Text_Count.Text = "回收站中有超过2,147,483,647个对象"
    Else
        Me.Text_Count.Text = "回收站中包含" + _
                        Str(rbinfo.i64NumItems.LowPart) + "个对象"
    End If
    '显示回收站中目前有多少对象
    
    If (rbinfo.i64Size.LowPart And &H80000000) = &H80000000 Or _
        rbinfo.i64Size.HighPart > 0 Then
        Me.Text_Space.Text = "回收站已用空间超过2,147,483,647字节"
    Else
        Me.Text_Space.Text = "回收站已用空间" + _
                            Str(rbinfo.i64Size.LowPart) + "字节"
    End If
    '显示回收站中的对象占了多少空间
End Sub

Private Sub Form_Load()
    Drive1_Change
End Sub
