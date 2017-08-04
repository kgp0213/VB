Attribute VB_Name = "Module1"
Public Declare Function CreatePolygonRgn _
Lib "gdi32" _
( _
    lpPoint As POINTAPI, ByVal nCount As Long, _
    ByVal nPolyFillMode As Long _
    ) As Long
Public Type POINTAPI
        x As Long
        y As Long
End Type
'------------------------------------------------
'以上为声明CreatePolygonRgn函数和它需要的POINTAPI类型

Public Declare Function SetWindowRgn _
Lib "user32" _
( _
    ByVal hWnd As Long, ByVal hRgn As Long, _
    ByVal bRedraw As Boolean _
    ) As Long
Public Const ALTERNATE = 1
'------------------------------------------
''以上为声明SetWindowRgn函数和它需要的常量

Public Sub Poly(f As Form)
    Dim hdc1 As Long
    Dim rec(37) As POINTAPI
    rec(0).x = 102
    rec(0).y = 11
    rec(1).x = 12
    rec(1).y = 102
    rec(2).x = 11
    rec(2).y = 107
    rec(3).x = 10
    rec(3).y = 119
    rec(4).x = 11
    rec(4).y = 124
    rec(5).x = 13
    rec(5).y = 127
    rec(6).x = 72
    rec(6).y = 187
    rec(7).x = 51
    rec(7).y = 211
    rec(8).x = 50
    rec(8).y = 227
    rec(9).x = 52
    rec(9).y = 230
    rec(10).x = 52
    rec(10).y = 233
    rec(11).x = 54
    rec(11).y = 235
    rec(12).x = 55
    rec(12).y = 237
    rec(13).x = 60
    rec(13).y = 241
    rec(14).x = 64
    rec(14).y = 234
    rec(15).x = 66
    rec(15).y = 246
    rec(16).x = 72
    rec(16).y = 248
    rec(17).x = 82
    rec(17).y = 248
    rec(18).x = 86
    rec(18).y = 247
    rec(19).x = 87
    rec(19).y = 247
    rec(20).x = 88
    rec(20).y = 246
    rec(21).x = 185
    rec(21).y = 150
    rec(22).x = 187
    rec(22).y = 145
    rec(23).x = 189
    rec(23).y = 138
    rec(24).x = 187
    rec(24).y = 129
    rec(25).x = 186
    rec(25).y = 125
    rec(26).x = 183
    rec(26).y = 121
    rec(27).x = 127
    rec(27).y = 64
    rec(28).x = 142
    rec(28).y = 43
    rec(29).x = 143
    rec(29).y = 30
    rec(30).x = 140
    rec(30).y = 24
    rec(31).x = 139
    rec(31).y = 21
    rec(32).x = 126
    rec(32).y = 10
    rec(33).x = 120
    rec(33).y = 9
    rec(34).x = 110
    rec(34).y = 9
    rec(35).x = 104
    rec(35).y = 12
    hdc1 = CreatePolygonRgn(rec(0), 37, ALTERNATE)
    SetWindowRgn f.hWnd, hdc1, True
End Sub


