VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   3255
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const COLOR_SCROLLBAR = 0 '滚动条
Private Const COLOR_BACKGROUND = 1 '桌面背景
Private Const COLOR_ACTIVECAPTION = 2 '活动窗口标题
Private Const COLOR_INACTIVECAPTION = 3 '非活动窗口标题
Private Const COLOR_MENU = 4 '菜单
Private Const COLOR_WINDOW = 5 '窗口背景
Private Const COLOR_WINDOWFRAME = 6 '窗口框
Private Const COLOR_MENUTEXT = 7 '窗口文字
Private Const COLOR_WINDOWTEXT = 8 '3D 阴影 (Win95)
Private Const COLOR_CAPTIONTEXT = 9 '标题文字
Private Const COLOR_ACTIVEBORDER = 10 '活动窗口边框
Private Const COLOR_INACTIVEBORDER = 11 '非活动窗口边框
Private Const COLOR_APPWORKSPACE = 12 'MDI窗口背景
Private Const COLOR_HIGHLIGHT = 13 '选择条背景
Private Const COLOR_HIGHLIGHTTEXT = 14 '选择条文字
Private Const COLOR_BTNFACE = 15 '按钮
Private Const COLOR_BTNSHADOW = 16 '3D 按钮阴影
Private Const COLOR_GRAYTEXT = 17 '灰度文字
Private Const COLOR_BTNTEXT = 18 '按钮文字
Private Const COLOR_INACTIVECAPTIONTEXT = 19 '非活动窗口文字
Private Const COLOR_BTNHIGHLIGHT = 20 '3D 选择按钮

Private Declare Function SetSysColors Lib "user32" _
                (ByVal nChanges As Long, _
                lpSysColor As Long, _
                lpColorValues As Long) _
                As Long
Private Declare Function GetSysColor Lib "user32" _
                (ByVal nIndex As Long) _
                As Long

Private Sub Form_Load()
    Dim i As Long
    i = GetSysColor(COLOR_ACTIVECAPTION)
    'i 是 RGB 值
    i = SetSysColors(1, COLOR_ACTIVECAPTION, RGB(255, 255, 255))
    '把标题设置为白色
End Sub
