Attribute VB_Name = "Module1"
Option Explicit
Public Sub Main()
     Dim x, y
     x = Screen.Width / Screen.TwipsPerPixelX
     y = Screen.Height / Screen.TwipsPerPixelY
     MsgBox "您的电脑解析度是" & x & " * " & y
     End
End Sub
     
