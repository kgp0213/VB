VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "动态装入菜单项"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu File 
      Caption         =   "文件"
      Index           =   0
      Begin VB.Menu Open 
         Caption         =   "打开"
         Index           =   0
      End
      Begin VB.Menu OpenedList 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filenamearry(3) As String
'声明数组用来存储最近打开的三个文件的名称
Private Sub Form_Load()
    Load OpenedList(1)
    Load OpenedList(2)
    '装入两个菜单项
    OpenedList(1).Visible = False
    OpenedList(2).Visible = False
    '设置这两个菜单项
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload OpenedList(1)
    Unload OpenedList(2)
    '从内存中卸载这两个菜单项
End Sub

Private Sub Open_Click(Index As Integer)
    CommonDialog1.ShowOpen
    '显示Open对话框
    filenamearry(2) = filenamearry(1)
    '将filenamearry(1)存储到filenamearry(2)
    filenamearry(1) = filenamearry(0)
    '将filenamearry(0)存储到filenamearry(1)
    filenamearry(0) = CommonDialog1.FileName
    '将此时打开的文件的名称存储到filenamearry(0)
    OpenedList(0).Caption = "&1" + filenamearry(0)
    OpenedList(0).Visible = True
    '显示刚刚打开的文件名称
    If Len(filenamearry(2)) > 0 Then
        OpenedList(2).Caption = "&3" + filenamearry(2)
        OpenedList(2).Visible = True
    End If
    If Len(filenamearry(1)) > 0 Then
        OpenedList(1).Caption = "&2" + filenamearry(1)
        OpenedList(1).Visible = True
    End If
End Sub

Private Sub OpenedList_Click(Index As Integer)
Select Case Index
    Case 0
        MsgBox "打开文件" + filenamearry(0), vbOKOnly, "打开"
    Case 1
        MsgBox "打开文件" + filenamearry(1), vbOKOnly, "打开"
    Case 2
        MsgBox "打开文件" + filenamearry(2), vbOKOnly, "打开"
End Select
End Sub
