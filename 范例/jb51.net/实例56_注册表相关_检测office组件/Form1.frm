VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "开始检测"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "你的计算机上有以下Office组件"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RegOpenKey Lib _
"advapi32" Alias "RegOpenKeyA" (ByVal hKey _
As Long, ByVal lpSubKey As String, _
phkResult As Long) As Long

Private Declare Function RegQueryValueEx _
Lib "advapi32" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As _
String, lpReserved As Long, lptype As _
Long, lpData As Any, lpcbData As Long) _
As Long

Private Declare Function RegCloseKey& Lib _
"advapi32" (ByVal hKey&)

Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const ERROR_SUCCESS = 0
Private Const HKEY_CLASSES_ROOT = &H80000000

Function GetRegString(hKey As Long, _
    strSubKey As String, strValueName As _
    String) As String
    Dim strSetting As String
    Dim lngDataLen As Long
    Dim lngRes As Long
    
    If RegOpenKey(hKey, strSubKey, _
        lngRes) = ERROR_SUCCESS Then
        strSetting = Space(255)
        lngDataLen = Len(strSetting)
        If RegQueryValueEx(lngRes, _
            strValueName, ByVal 0, _
            REG_EXPAND_SZ, ByVal strSetting, _
            lngDataLen) = ERROR_SUCCESS Then
                If lngDataLen > 1 Then
                    GetRegString = Left(strSetting, _
                    lngDataLen - 1)
                End If
        End If

        If RegCloseKey(lngRes) <> ERROR_SUCCESS Then
            MsgBox "RegCloseKey Failed: " & _
            strSubKey, vbCritical
        End If
    End If
End Function
Function FileExists(sFileName$) As Boolean
    On Error Resume Next
    FileExists = IIf(Dir(Trim(sFileName)) <> "", _
    True, False)
End Function

Public Function IsAppPresent(strSubKey$, _
    strValueName$) As Boolean
    IsAppPresent = CBool(Len(GetRegString(HKEY_CLASSES_ROOT, _
    strSubKey, strValueName)))
End Function

Private Sub Command1_Click()
    Label1.Caption = "Access " & _
        IsAppPresent("Access.Database\CurVer", "")

    Label2.Caption = "Excel " & _
        IsAppPresent("Excel.Sheet\CurVer", "")

    Label3.Caption = "PowerPoint " & _
        IsAppPresent("PowerPoint.Slide\CurVer", "")

    Label4.Caption = "Word " & _
        IsAppPresent("Word.Document\CurVer", "")
End Sub

Private Sub Form_Load()

End Sub
