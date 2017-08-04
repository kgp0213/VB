VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Message"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   2535
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SubClass(wnd As Long)
    Dim ret As Long
    '记录Window Procedure的地址
    preWinProc = GetWindowLong(wnd, GWL_WNDPROC)
    '开始截取消息,并将消息交给wndproc过程处理.
    ret = SetWindowLong(wnd, GWL_WNDPROC, AddressOf wndproc)
End Sub

Private Sub EndSubClass(wnd As Long)
    Dim ret As Long
    '取消消息截取，结束子分类过程.
    ret = SetWindowLong(wnd, GWL_WNDPROC, preWinProc)
End Sub

Private Sub Form_Load()
    SubClass Me.Text1.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndSubClass Me.Text1.hwnd
End Sub
