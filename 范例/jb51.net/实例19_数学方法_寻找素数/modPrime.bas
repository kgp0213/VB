Attribute VB_Name = "modPrime"
Public lngStart As Long
Public lngEnd As Long
Public lngCount As Long

Function PrimeStatus(TestVal As Long) As Boolean
   Dim Lim As Integer
   PrimeStatus = True
   Lim = Sqr(TestVal)
   For I = 2 To Lim
      If TestVal Mod I = 0 Then
         PrimeStatus = False
         Exit For
      End If
      'If I Mod 200 = 0 Then DoEvents
   Next I
End Function

