VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "椭圆形窗口"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4080
   ScaleWidth      =   4860
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateEllipticRgn Lib "gdi32" _
( _
    ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long _
    ) As Long
'声明创建椭圆形区域的API函数
Private Declare Function CreateRectRgn Lib "gdi32" _
( _
    ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long _
    ) As Long
'声明创建矩形区域的API函数
'该函数将用来将窗口恢复为矩形

Private Declare Function SetWindowRgn Lib "user32" _
( _
    ByVal hWnd As Long, ByVal hRgn As Long, _
    ByVal bRedraw As Boolean _
    ) As Long
'声明设置窗口形状的API函数
Dim hRgnC As Long
Dim hRgnR As Long
'声明变量用来存储椭圆形区域和矩形区域的句柄


Private Sub Form_Click()
    'hRgnC = CreateEllipticRgn(10, 10, 200, 200)
    'hRgnC = CreateEllipticRgn(0, 0, 150, 200)
    hRgnC = CreateEllipticRgn(20, 0, 500, 500)
    '创建椭圆形区域
    SetWindowRgn Me.hWnd, hRgnC, True
    '设置窗口为椭圆形
End Sub

Private Sub Form_DblClick()
    hRgnR = CreateRectRgn(0, 0, Me.Width, Me.Height)
    '创建矩形区域
    SetWindowRgn Me.hWnd, hRgnR, True
    '设置窗口为矩形
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If hRgnR <> 0 Then DeleteObject hRgnR
    If hRgnC <> 0 Then DeleteObject hRgnC
    '释放创建的图形区域
End Sub
