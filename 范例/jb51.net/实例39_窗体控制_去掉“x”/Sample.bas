Attribute VB_Name = "Module1"
Option Explicit
'第一种方法
Declare Function GetSystemMenu Lib "User32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu As Long) As Long
Public Const MF_BYPOSITION = &H400&

'第二种方法
'Declare Function GetSystemMenu Lib "User32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function RemoveMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DrawMenuBar Lib "User32" (ByVal hwnd As Long) As Long
'Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu As Long) As Long
'Public Const MF_BYPOSITION = &H400&
Public Const MF_DISABLED = &H2&

'第一种方法
Public Sub DisableX(Frm As Form)
    Dim hMenu As Long, nCount As Long
    hMenu = GetSystemMenu(Frm.hwnd, 0)
    nCount = GetMenuItemCount(hMenu)
    Call RemoveMenu(hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION)
    DrawMenuBar Frm.hwnd
End Sub
