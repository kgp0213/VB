Attribute VB_Name = "Module1"
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal Wid As Long, ByVal Heit As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hobject As Long) As Long

