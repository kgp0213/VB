VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Icon"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   3630
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.DirListBox Dir1 
      Height          =   1980
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更改图标"
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4

Dim sPath As String
'变量sPath用来存储要更改图标的文件夹的完整路径

Private Sub Form_Load()
    CommonDialog1.Filter = "(*.ico)|*.ico"
    CommonDialog1.DialogTitle = "选择图标"
    Drive1.Drive = "c:\"
    Dir1.Path = "c:\inetpub"
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
    Dim i As Integer
    i = Dir1.ListIndex
    'ListIndex属性返回或者设置当前选中文件夹的索引
    sPath = Dir1.List(i)
    'sPath存储文件夹的完整路径
End Sub

Private Sub Command1_Click()
    ChangeFolderInfo sPath
    '更改当前选中的文件夹为系统文件
    
    Dim s As String
    '声明变量s用来保存图标文件路径
    CommonDialog1.ShowOpen
    '显示“打开”对话框
    s = CommonDialog1.FileName
    '将在对话框中选中的文件的完整路径保存在s中
    If s = "" Then Exit Sub
    Dim sFile As String
    If Right(sPath, 1) = "\" Then
        sFile = sPath + "desktop.ini"
    Else
        sFile = sPath + "\" + "desktop.ini"
    End If
    
    Dim num_File As Integer
    num_File = FreeFile()
    If Dir(sFile, vbHidden) <> "" Then
        Call ChangeFileInfo(sFile, 0)
    End If
    
    Open sFile For Output As #num_File
    '打开目前选中的文件夹的desktop.ini文件
    Print #num_File, "[.ShellClassInfo]" + vbCrLf + "IconIndex=0" + _
                     vbCrLf + "iconfile=" + s
    '更改desktop.ini文件
    Close #num_File
    Call ChangeFileInfo(sFile, FILE_ATTRIBUTE_HIDDEN)
    '将desktop.ini文件设置为隐藏
End Sub

Private Sub ChangeFolderInfo(folderspec)
'该子过程赋予文件夹系统属性
    If Len(folderspec) < 1 Then
        Exit Sub
    End If
    
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    f.Attributes = FILE_ATTRIBUTE_SYSTEM
    '设置文件夹属性
End Sub

Private Sub ChangeFileInfo(filespec, attr)
'该子过程赋予Desktop.ini文件隐藏属性
    If Len(filespec) < 1 Then
        Exit Sub
    End If
    
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    f.Attributes = attr
    '设置文件属性
End Sub

