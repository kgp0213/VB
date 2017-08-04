VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "抓出应用程式的图标并存档"
   ClientHeight    =   1995
   ClientLeft      =   2805
   ClientTop       =   2490
   ClientWidth     =   4815
   Icon            =   "ExtractIcon.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   1995
   ScaleWidth      =   4815
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   1410
      TabIndex        =   5
      Top             =   600
      Width           =   3240
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   570
         Width           =   3000
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   3000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   150
      TabIndex        =   2
      Top             =   600
      Width           =   1215
      Begin VB.VScrollBar VScroll1 
         Height          =   615
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   630
         Left            =   120
         ScaleHeight     =   570
         ScaleWidth      =   570
         TabIndex        =   3
         ToolTipText     =   "Use the scroll bar to view the images"
         Top             =   240
         Width           =   630
      End
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   390
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "开启档案"
      Height          =   360
      Left            =   1470
      TabIndex        =   1
      ToolTipText     =   "Select a different resource"
      Top             =   180
      Width           =   1200
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "另存图示"
      Height          =   360
      Left            =   150
      TabIndex        =   0
      ToolTipText     =   "Save the currently selected image into a file"
      Top             =   180
      Width           =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   210
      TabIndex        =   8
      Top             =   1650
      Width           =   4560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lIcon As Long
Dim sSourcePgm As String
Dim sDestFile As String

Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" _
(ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Private Declare Function DrawIcon Lib "user32" _
(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Sub CmdSave_Click() '另存图示
  On Error Resume Next
  With Dlg                  '存档问话框
    .FileName = sDestFile
    .CancelError = True
    .Action = 2
    If Err Then
      Err.Clear
      Exit Sub
    End If
    sDestFile = .FileName
    SavePicture Picture1.Image, sDestFile '将抓出的图示存档
  End With
End Sub

Private Sub CmdOpen_Click() '开启档案
  Dim a%
  
  On Error Resume Next
  With Dlg                  '开档问话框
    .FileName = sSourcePgm
    .CancelError = True
    .DialogTitle = "请选择包含图示的 DLL 或 EXE 档"
    .Filter = "Icon Resources (*.ico;*.exe;*.dll)|*.ico;*.exe;*.dll|All files|*.*"
    .Action = 1
    If Err Then
      Err.Clear
      Exit Sub
    End If
    sSourcePgm = .FileName
    Label3.Caption = .FileName
    DestroyIcon lIcon
    Do
      lIcon = ExtractIcon(App.hInstance, sSourcePgm, a)
      If lIcon = 0 Then Exit Do
      a = a + 1
      DestroyIcon lIcon
    Loop
    If a = 0 Then
      MsgBox "在这个档中没有任何图示！"
    End If
    Label1.Caption = "在这个档中共有 " & a & " 个图示"
    VScroll1.Max = IIf(a = 0, 0, a - 1)
    VScroll1.Value = 0
    VScroll1_Change
  End With
End Sub

Private Sub Form_Load()
  CmdOpen_Click
End Sub


Private Sub Picture1_Click()

End Sub

Private Sub VScroll1_Change()
  Label2.Caption = "正在浏览的图示索引值： " & VScroll1.Value
  DestroyIcon lIcon
  Picture1.Cls
  lIcon = ExtractIcon(App.hInstance, sSourcePgm, VScroll1.Value)
  Picture1.AutoSize = True
  Picture1.AutoRedraw = True
  DrawIcon Picture1.hdc, 0, 0, lIcon
  Picture1.Refresh
End Sub
