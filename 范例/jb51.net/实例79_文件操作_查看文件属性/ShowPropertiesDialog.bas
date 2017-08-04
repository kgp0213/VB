Attribute VB_Name = "ShowPropertiesDialog"
Option Explicit

Type SHELLEXECUTEINFO
    cbSize        As Long
    fMask         As Long
    hwnd          As Long
    lpVerb        As String
    lpFile        As String
    lpParameters  As String
    lpDirectory   As String
    nShow         As Long
    hInstApp      As Long
    
    lpIDList      As Long     ' Optional parameter
    lpClass       As String   ' Optional parameter
    hkeyClass     As Long     ' Optional parameter
    dwHotKey      As Long     ' Optional parameter
    hIcon         As Long     ' Optional parameter
    hProcess      As Long     ' Optional parameter
End Type

Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400

Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" _
(SEI As SHELLEXECUTEINFO) As Long
Public Function ShowProperties(FileName As String, OwnerhWnd As Long) As Long
  
    ' Open a file properties dialog for specified file if return value
    ' <=32 an error occurred

    Dim SEI As SHELLEXECUTEINFO
    Dim r As Long
     
    ' Fill in the SHELLEXECUTEINFO structure
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = FileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
     
    ' Call the API
    r = ShellExecuteEX(SEI)
     
    ' Return the instance handle as a sign of success
    ShowProperties = SEI.hInstApp
  
End Function

