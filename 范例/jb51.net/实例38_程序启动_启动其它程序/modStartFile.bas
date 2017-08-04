Attribute VB_Name = "modStartFile"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Function StartDoc(DocName As String) As Long
  Dim Scr_hDC As Long
  Scr_hDC = GetDesktopWindow()
  
  'change "Open" to "Explore" to bring up file explorer
  StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", 1)
End Function
