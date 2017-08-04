VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileInt Lib "Kernel32" _
                Alias "GetPrivateProfileIntA" _
                (ByVal lpApplicationName As String, _
                ByVal lpKeyName As String, _
                ByVal nDefault As Long, _
                ByVal lpFileName As String) _
                As Long
Private Declare Function GetPrivateProfileString Lib "Kernel32" _
                Alias "GetPrivateProfileStringA" _
                (ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, _
                ByVal lpDefault As String, _
                ByVal lpReturnedString As String, _
                ByVal nSize As Integer, _
                ByVal lpFileName As String) _
                As Integer

Private Declare Function WritePrivateProfileString Lib "Kernel32" _
                Alias "WritePrivateProfileStringA" _
                (ByVal lpApplicationName As String, _
                ByVal lpKeyName As String, _
                ByVal lpString As String, _
                ByVal lpFileName As String) _
                As Long
Private Declare Function GetPrivateProfileSection Lib "Kernel32" _
                Alias "GetPrivateProfileSectionA" _
                (ByVal lpAppName As String, _
                ByVal lpReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpFileName As String) _
                As Long
Private Declare Function WritePrivateProfileSection Lib "Kernel32" _
                Alias "WritePrivateProfileSectionA" _
                (ByVal lpAppName As String, _
                ByVal lpString As String, _
                ByVal lpFileName As String) _
                As Long

Dim FILE_NAME As String

Private Sub Form_Load()
    FILE_NAME = App.Path + "\test.ini"
    Dim strCaption As String * 256
    Call GetPrivateProfileString("Form", "Caption", "Default Caption", _
                        strCaption, 256, FILE_NAME)
    Me.Caption = Trim(strCaption)
    Me.Width = GetPrivateProfileInt("Form", "Width", Me.Width, FILE_NAME)
    Me.Height = GetPrivateProfileInt("Form", "Height", Me.Height, FILE_NAME)
    Me.Left = GetPrivateProfileInt("Form", "Left", Me.Left, FILE_NAME)
    Me.Top = GetPrivateProfileInt("Form", "Top", Me.Top, FILE_NAME)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strCaption As String
    strCaption = Me.Caption
    Call WritePrivateProfileString("Form", "Caption", strCaption, FILE_NAME)
    Call WritePrivateProfileString("Form", "Width", Str(Me.Width), FILE_NAME)
    Call WritePrivateProfileString("Form", "Height", Str(Me.Height), FILE_NAME)
    Call WritePrivateProfileString("Form", "Left", Str(Me.Left), FILE_NAME)
    Call WritePrivateProfileString("Form", "Top", Str(Me.Top), FILE_NAME)
End Sub
