Attribute VB_Name = "Module1"
Public Declare Function GetWindowLong _
Lib "user32" Alias "GetWindowLongA" _
( _
    ByVal hWnd As Long, ByVal nIndex As Long _
 ) As Long
'hwnd为窗口句柄
'nIndex指示要获得窗口哪方面的特征
'nIndex参数可以为下列常量之一：
'GWL_EXSTYLE
'GWL_HINSTANCE
'GWL_HWNDPARENT
'GWL_ID
'GWL_STYLE
'GWL_WNDPROC
'GWL_USERDATA
'---------------------------------------------
Public Declare Function SetWindowLong _
Lib "user32" Alias "SetWindowLongA" _
( _
    ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
 ) As Long
'hwnd为要设置特征的窗口的句柄
'nIndex指示要设置窗口哪方面特征
'dwNewLong为表示窗口信息的一个Long类型数值
'---------------------------------------------
Public Const GWL_STYLE = (-16)
Public Const WS_CAPTION = &HC00000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000
'声明SetWindowLong和GetWindowLong函数将要使用的常量
'---------------------------------------------

Public Declare Function SetWindowPos Lib "user32" _
( _
    ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long _
) As Long

Public Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20
    'The frame changed: send WM_NCCALCSIZE
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    'Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum
'---------------------------------------------
Public Declare Function GetWindowRect Lib "user32" _
( _
    ByVal hWnd As Long, lpRect As RECT _
    ) As Long
'GetWindowRect函数获得整个窗口的范围矩形
'窗口的边框、标题栏、滚动条及菜单等都在这个矩形内
'hWnd参数为Long型，要获得范围矩形的窗口的句柄
'lpRect参数为RECT结构，屏幕坐标中随同窗口装载的矩形
   
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'---------------------------------------------
