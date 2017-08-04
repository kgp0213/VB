Attribute VB_Name = "Module1"
Public Declare Function AlphaBlend Lib "msimg32" _
( _
    ByVal hDestDC As Long, _
    ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal widthSrc As Long, ByVal heightSrc As Long, _
    ByVal blendFunct As Long _
    ) As Boolean

Public Type rBlendProps
    tBlendOp As Byte
    tBlendOptions As Byte
    tBlendAmount As Byte
    tAlphaType As Byte
End Type


Public Declare Sub CopyMemory Lib "kernel32" _
Alias "RtlMoveMemory" _
( _
    Destination As Any, Source As Any, _
    ByVal Length As Long)


