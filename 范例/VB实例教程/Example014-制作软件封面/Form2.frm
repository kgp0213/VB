VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "主窗口"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form2"
   ScaleHeight     =   4275
   ScaleWidth      =   5475
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   720
      Top             =   600
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Timer1.Interval = 3000
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Form1.Show
    SetWindowPos Form1.hwnd, HWND_TOPMOST, _
        0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    '鼠标呈沙漏状
    '显示封面
 End Sub

Private Sub Timer1_Timer()
    Form2.Timer1.Enabled = False
    '关闭定时器
    Unload Form1
    '卸载封面
    Screen.MousePointer = 0
    '鼠标恢复原样
End Sub
