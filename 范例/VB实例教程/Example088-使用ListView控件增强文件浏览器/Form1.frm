VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "文件浏览器"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   7710
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   5130
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   1350
      Left            =   2400
      Pattern         =   "*.exe"
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5775
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   10186
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clmX As ColumnHeader    '标题栏
Dim itmX As ListItem        '列表项目
Dim Counter As Long         '计数器
Dim Fname As String         '读取文件名
Dim CurrentDir As String
Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    '窗体位置在屏幕中间
    '以下代码为ListView1添加标题栏
    ListView1.ColumnHeaders.Add , , "文件名称", ListView1.Width / 3, 0
    '第一个标题栏是"文件名称",长度为ListView1宽度的三分之一 , 文字左对齐
    Set clmX = ListView1.ColumnHeaders.Add(, , "序号", ListView1.Width / 6, 2)
    '第二个标题栏是"序号"
    Set clmX = ListView1.ColumnHeaders.Add(, , "文件大小", ListView1.Width / 4, 1)
    '第三个标题栏是"文件大小"
    Set clmX = ListView1.ColumnHeaders.Add(, , "创建时间", ListView1.Width / 3, 0)
    '第四个标题栏是"创建时间"
    ListView1.SmallIcons = ImageList1
    '关联ImageList1中的图标
    
    For Counter = 0 To File1.ListCount - 1
        Fname = File1.List(Counter)
        Set itmX = ListView1.ListItems.Add(, , Fname)
        '添加文件名
        itmX.SubItems(1) = CStr(Counter + 1) + "/" + CStr(File1.ListCount)
        itmX.SubItems(2) = CStr(FileLen(CurrentDir & Fname))
        itmX.SmallIcon = 1
        itmX.SubItems(3) = Format(FileDateTime(CurrentDir + Fname), "HH:MM YYYY/MMMM/DD")
     Next Counter
      '添加ListView的各个项目
    
End Sub

Private Static Sub Drive1_Change()
    On Error GoTo IFerr '拦截错误
    Dir1.Path = Drive1.Drive
    '关联目录列表框
    Exit Sub
IFerr:                 '如果磁盘错误
    MsgBox ("请确认驱动器是否准备好或者磁盘已经不可用!"), vbOKOnly + vbExclamation
    '弹出注意对话框
    Drive1.Drive = Dir1.Path
    '忽略驱动器改变
End Sub

Private Static Sub Dir1_Change()
    File1.Path = Dir1.Path
    '关联文件列表框
    If Right(Dir1.Path, 1) <> "\" Then
        CurrentDir = Dir1.Path & "\"
    Else
        CurrentDir = Dir1.Path
    End If
    '设置选定的目录名称
    ListView1.ListItems.Clear
    '清除过期的列表项目
    For Counter = 0 To File1.ListCount - 1
        Fname = File1.List(Counter)
        Set itmX = ListView1.ListItems.Add(, , Fname)
        '添加文件名
        itmX.SubItems(1) = CStr(Counter + 1) + "/" + CStr(File1.ListCount)
        itmX.SubItems(2) = CStr(FileLen(CurrentDir & Fname))
        itmX.SmallIcon = 1
        itmX.SubItems(3) = Format(FileDateTime(CurrentDir + Fname), "HH:MM YYYY/MMMM/DD")
     Next Counter
      '添加ListView的各个项目
End Sub
