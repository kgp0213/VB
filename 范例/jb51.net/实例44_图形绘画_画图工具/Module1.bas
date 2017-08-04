Attribute VB_Name = "Module1"
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal nxpos As Long, ByVal nypos As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal colorref As Long) As Long
