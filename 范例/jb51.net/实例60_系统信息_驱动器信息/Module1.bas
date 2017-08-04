Attribute VB_Name = "Module1"

Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Global Const DRIVE_REMOVABLE = 2
Global Const DRIVE_FIXED = 3
Global Const DRIVE_REMOTE = 4
Global Const DRIVE_CDROM = 5
Global Const DRIVE_RAMDISK = 6

Public Function GetAllDrives() As String
Dim lngResult&, strDrives$, strJustOneDrive$, intPos%, lngDriveType&
   
   strDrives$ = Space$(255)
   lngResult& = GetLogicalDriveStrings(Len(strDrives$), strDrives$)
  
   strDrives$ = Left$(strDrives$, lngResult&)
   
    Do
      intPos% = InStr(strDrives$, Chr$(0))
      If intPos% Then
        strJustOneDrive$ = Left$(strDrives$, intPos%)
        strDrives$ = Mid$(strDrives$, intPos% + 1, Len(strDrives$))
        lngDriveType& = GetDriveType(strJustOneDrive$)
        
        Select Case lngDriveType&
        Case 5
          strBuffer = strBuffer & "CD-Rom: " & strJustOneDrive$ & vbCrLf
        Case 2
          strBuffer = strBuffer & "RemovableDrive: " & strJustOneDrive$ & vbCrLf
        Case 3
          strBuffer = strBuffer & "LocalDrive: " & strJustOneDrive$ & vbCrLf
        Case 4
          strBuffer = strBuffer & "NetworkDrive: " & strJustOneDrive$ & vbCrLf
        Case 6
          strBuffer = strBuffer & "RamDrive: " & strJustOneDrive$ & vbCrLf
       End Select
      
      End If
   
    Loop Until strDrives$ = ""
    GetAllDrives = strBuffer
    
End Function

