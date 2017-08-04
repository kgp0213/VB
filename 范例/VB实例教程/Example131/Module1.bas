Attribute VB_Name = "Module1"
Public Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
                (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) _
                As Long

Public Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
                (ByVal hwnd As Long, _
                ByVal nIndex As Long) _
                As Long

Public Declare Function CallWindowProc Lib "user32" _
                Alias "CallWindowProcA" _
                (ByVal lpPrevWndFunc As Long, _
                ByVal hwnd As Long, _
                ByVal Msg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) _
                As Long
'定义常数
Public Const GWL_WNDPROC = (-4)
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

'全局变量,存放控件标志性数据
Public preWinProc As Long

'本函数就是用来接收子类时截取的消息的
Public Function wndproc(ByVal hwnd As Long, ByVal Msg As Long, _
                ByVal wParam As Long, ByVal lParam As Long) As Long

    '截取下来的消息存放在msg参数中.
    If (Msg = WM_RBUTTONDOWN) Or (Msg = WM_RBUTTONUP) Then
        '检测到消息,这里就可以加入处理代码
        '需要注意,如果这儿不加入任何代码,则相当于吃掉了这条消息.
    Else
        '如果不是需要处理的消息,则将之送回原来的程序.
        wndproc = CallWindowProc(preWinProc, hwnd, Msg, wParam, lParam)
    End If
End Function
