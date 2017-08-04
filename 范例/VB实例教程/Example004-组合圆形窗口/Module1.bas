Attribute VB_Name = "Module1"
Public Declare Function CreateEllipticRgn Lib "gdi32" _
( _
    ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long _
    ) As Long

Public Declare Function CreateRectRgn Lib "gdi32" _
( _
    ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long _
    ) As Long

Public Declare Function SetWindowRgn Lib "user32" _
( _
    ByVal hWnd As Long, ByVal hRgn As Long, _
    ByVal bRedraw As Boolean _
    ) As Long

Public Declare Function CombineRgn Lib "gdi32" _
( _
    ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, _
    ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long _
    ) As Long

Public Const RGN_AND = 1
'将两个区域相加
Public Const RGN_COPY = 5
'创建hSrcRgn1的拷贝
Public Const RGN_DIFF = 4
'将两个区域相减
Public Const RGN_OR = 2
'将两个区域进行或操作
Public Const RGN_XOR = 3
'将两个区域进行异或操作
Public Const RGN_MAX = RGN_COPY
Public Const RGN_MIN = RGN_AND

Public Sub SetWindow(f1 As Form)
'该子过程实现设置窗口形状
    Dim hSrcRgn1, hSrcRgn2, hSrcRgn3 As Long
    hSrcRgn1 = CreateEllipticRgn(5, 23, 397, 415)
    '创建最外面的大圆区域
    hSrcRgn2 = CreateEllipticRgn(90, 74, 395, 362)
    '创建中间的圆区域
    hSrcRgn3 = CreateEllipticRgn(183, 120, 395, 320)
    '创建最里层的小圆区域
    
    CombineRgn hSrcRgn1, hSrcRgn1, hSrcRgn2, RGN_DIFF
    '用大圆减去中间的圆得到的区域保存在hSrcRgn1
    CombineRgn hSrcRgn1, hSrcRgn1, hSrcRgn3, RGN_OR
    '用得到的区域加上小圆并保存在hSrcRgn1
    SetWindowRgn f1.hWnd, hSrcRgn1, True
End Sub
Public Sub Reset(f1 As Form)
'该子过程实现恢复窗口形状为矩形
    Dim hSrcRgn4 As Long
    hSrcRgn4 = CreateRectRgn(0, 0, f1.Width, f1.Height)
    '创建矩形
    SetWindowRgn f1.hWnd, hSrcRgn4, True
    '将窗口恢复为矩形
End Sub
