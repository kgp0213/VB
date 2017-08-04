Attribute VB_Name = "Module1"
Public Declare Function GetMenu Lib "user32" _
( _
    ByVal hwnd As Long _
    ) As Long
'获得窗口中的菜单
'-------------------------------------------------------
Public Declare Function GetSubMenu Lib "user32" _
( _
    ByVal hMenu As Long, _
    ByVal nPos As Long _
    ) As Long
'获得菜单中的子菜单(下一级菜单)
'-------------------------------------------------------
Public Declare Function GetMenuItemID Lib "user32" _
( _
    ByVal hMenu As Long, _
    ByVal nPos As Long _
    ) As Long
'获得菜单项
'-------------------------------------------------------
Public Declare Function GetMenuItemCount Lib "user32" _
( _
    ByVal hMenu As Long _
    ) As Long
'获得指定菜单下菜单项的数目
'-------------------------------------------------------
Public Declare Function GetMenuItemInfo Lib "user32" _
Alias "GetMenuItemInfoA" _
( _
    ByVal hMenu As Long, ByVal un As Long, _
    ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO _
    ) As Boolean
'获得指定菜单项的信息
Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_STRING = &H0&
'-------------------------------------------------------
Public Declare Function SetMenuItemBitmaps Lib "user32" _
( _
    ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long, _
    ByVal hBitmapUnchecked As Long, _
    ByVal hBitmapChecked As Long _
    ) As Long
'设置菜单项
Public Const MF_BITMAP = &H4&
'-------------------------------------------------------

