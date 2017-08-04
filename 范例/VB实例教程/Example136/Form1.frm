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
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_USERS = &H80000003

Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const REG_DWORD_LITTLE_ENDIAN = 4
'32-bit number (same as REG_DWORD)
Private Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Private Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Private Const REG_LINK = 6                       ' Symbolic Link (unicode)
Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Private Const REG_NONE = 0                       ' No value type
Private Const REG_SZ = 1                         ' Unicode nul terminated string

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
                Alias "RegOpenKeyExA" _
                (ByVal hKey As Long, _
                ByVal lpSubKey As String, _
                ByVal ulOptions As Long, _
                ByVal samDesired As Long, _
                phkResult As Long) _
                As Long
                
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
                Alias "RegQueryValueExA" _
                (ByVal hKey As Long, _
                ByVal lpValueName As String, _
                ByVal lpReserved As Long, _
                lpType As Long, _
                lpData As Any, _
                lpcbData As Long) _
                As Long
                
Private Declare Function RegSetValueEx Lib "advapi32.dll" _
                Alias "RegSetValueExA" _
                (ByVal hKey As Long, _
                ByVal lpValueName As String, _
                ByVal Reserved As Long, _
                ByVal dwType As Long, _
                lpData As Any, _
                ByVal cbData As Long) _
                As Long

Private Declare Function RegCreateKey Lib "advapi32.dll" _
                Alias "RegCreateKeyA" _
                (ByVal hKey As Long, _
                ByVal lpSubKey As String, _
                phkResult As Long) _
                As Long
                
Private Declare Function RegDeleteValue Lib "advapi32.dll" _
                Alias "RegDeleteValueA" _
                (ByVal hKey As Long, _
                ByVal lpValueName As String) _
                As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" _
                (ByVal hKey As Long) _
                As Long
                
Const KEY_ALL_ACCESS = &HF3F 'Full access permission

Private Sub Form_Load()
        Dim handle As Long
        Dim ExecName As String
        Dim KeyName As String
        KeyName = "Software\Microsoft\Internet Explorer\Extensions\" + _
                    "{DA80DAD2-0220-448b-BA30-C84D5C678150}"

        Dim IconPathName As String      '正常时的图标全路径
        Dim HotIconPathName As String   '鼠标覆盖时的图标全路径
        
        If Right(App.Path, 1) = "\" Then
            ExecName = App.Path + App.EXEName + ".exe"
        Else
            ExecName = App.Path + "\" + App.EXEName + ".exe"
        End If
        
        IconPathName = ExecName + ",101"
        HotIconPathName = ExecName + ",102"
        
        Call RegCreateKey(HKEY_LOCAL_MACHINE, KeyName, handle)
        Dim str_Value As String
        str_Value = "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}"
        Call RegSetValueEx(handle, "CLSID", 0&, REG_SZ, ByVal str_Value, _
                            Len(str_Value) + 1)
        str_Value = "Yes"
        Call RegSetValueEx(handle, "Default Visible", 0&, REG_SZ, _
                            ByVal str_Value, Len(str_Value) + 1)
        str_Value = "Example"
        Call RegSetValueEx(handle, "ButtonText", 0&, REG_SZ, _
                            ByVal str_Value, Len(str_Value) + 1)
        Call RegSetValueEx(handle, "Icon", 0&, REG_SZ, ByVal IconPathName, _
                            LenB(StrConv(IconPathName, vbFromUnicode)))
        Call RegSetValueEx(handle, "HotIcon", 0&, REG_SZ, ByVal HotIconPathName, _
                            LenB(StrConv(HotIconPathName, vbFromUnicode)))
        'str_Value = "http://localhost"
        str_Value = "%SystemRoot%\Web\Example.htm"
        Call RegSetValueEx(handle, "Exec", 0&, REG_SZ, ByVal str_Value, _
                         LenB(StrConv(str_Value, vbFromUnicode)))
        RegCloseKey (handle)
End Sub
