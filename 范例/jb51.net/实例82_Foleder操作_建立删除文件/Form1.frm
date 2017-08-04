VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3836
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fsoSys As New Scripting.FileSystemObject
Dim fsoRootFolder As Folder

Private Sub Form_Load()
    Dim fsoSubFolder As Folder
    Dim nodRootNode As Node
    Dim nodChild As Node
    Dim astr$
    
    Set nodRootNode = TreeView1.Nodes.Add(, , "Root", "c:\")
    Set fsoRootFolder = fsoSys.GetFolder("c:\")
    For Each fsoSubFolder In fsoRootFolder.SubFolders
        astr = fsoSubFolder.Path
        Set nodChild = TreeView1.Nodes.Add("Root", tvwChild, astr, fsoSubFolder.Name)
    Next
    
    Set fsoRootFolder = Nothing
    Command1.Caption = "建立目录"
    Command2.Caption = "删除目录"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fsoSys = Nothing
End Sub

Private Sub Command1_Click()
    Dim fsoFolder As Folder
    
    '检查目录是否存在,如果目录不存在则建立新目录
    If fsoSys.FolderExists("c:\test") Then
        MsgBox ("目录c:\test已经存在，无法建立目录")
    Else
        Set fsoFolder = fsoSys.CreateFolder("c:\test")
        Set fsoFolder = Nothing
    End If
End Sub

Private Sub Command2_Click()
    '检查目录是否存在,如存在则删除目录
    If fsoSys.FolderExists("c:\test") Then
        fsoSys.DeleteFolder ("c:\test")
    Else
        MsgBox ("目录c:\test不存在")
    End If
End Sub

