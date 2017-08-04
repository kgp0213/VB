Attribute VB_Name = "Module1"
Type ULARGE_INTEGER
  LowPart As Long
  HighPart As Long
End Type

Type SHQUERYRBINFO
  cbSize As Long
  i64Size As ULARGE_INTEGER
  i64NumItems As ULARGE_INTEGER
End Type

Declare Function SHQueryRecycleBin Lib "shell32.dll" Alias "SHQueryRecycleBinA" _
(ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long

Public Const SHERB_NOCONFIRMATION = &H1
Public Const SHERB_NOPROGRESSUI = &H2
Public Const SHERB_NOSOUND = &H4

Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" _
(ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long

Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long


