Attribute VB_Name = "Module1"
Option Explicit
   
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_USERS = &H80000003

Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
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

Sub Main()
'程序运行入口
    Dim ExeCmd As String
    '存放命令行参数
    ExeCmd = UCase(Command$)
    '参数转换成大写后存放在变量ExeCmd里
    If Len(ExeCmd) > 0 Then
        Dim fso As Object
        Dim fso_file As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set fso_file = fso.OpenTextFile(ExeCmd, ForReading, True)
        Form1.Caption = ExeCmd
        Form1.Text1.Text = ""
        Do While Not fso_file.AtEndOfStream
            Form1.Text1.Text = Form1.Text1.Text + fso_file.ReadLine + Chr(13) + Chr(10)
        Loop
        fso_file.Close
        Form1.Show
    Else
        Dim handle As Long
        
        Dim FileExt As String
        Dim FileType As String
        Dim MIMEType As String
        Dim ExecName As String
        
        FileExt = ".test"
        FileType = "Test file"
        
        MIMEType = "text/plain"
        If Right(App.Path, 1) = "\" Then
            ExecName = App.Path + App.EXEName + ".exe %1"
        Else
            ExecName = App.Path + "\" + App.EXEName + ".exe %1"
        End If
        Call RegCreateKey(HKEY_CLASSES_ROOT, FileExt, handle)
        Call RegSetValueEx(handle, "", 0&, REG_SZ, ByVal FileType, Len(FileType) + 1)
        Call RegSetValueEx(handle, "Content Type", 0&, REG_SZ, ByVal MIMEType, _
                           Len(MIMEType) + 1)
        
        
        Call RegCreateKey(HKEY_CLASSES_ROOT, FileExt + "\shell\open\command", handle)
        Call RegSetValueEx(handle, "", 0&, REG_SZ, ByVal ExecName, LenB(StrConv(ExecName, vbFromUnicode)))
        RegCloseKey (handle)
    End If
End Sub
