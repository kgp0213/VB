VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "属性"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   3015
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "属性:"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2775
      Begin VB.CheckBox Check_System 
         Caption         =   "系统"
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Check_Archive 
         Caption         =   "存档"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Check_Hidden 
         Caption         =   "隐藏"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox Check_ReadOnly 
         Caption         =   "只读"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label_Location 
      AutoSize        =   -1  'True
      Caption         =   "位置:"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   450
   End
   Begin VB.Label Label_Name 
      AutoSize        =   -1  'True
      Caption         =   "名称:"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetFileAttributes Lib "kernel32" Alias _
        "GetFileAttributesA" (ByVal lpFileName As String) As Long

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4

Private Sub Form_Load()
    Me.CommonDialog1.ShowOpen
    If Me.CommonDialog1.FileName <> "" Then
        Dim str_File As String
        Dim attr As Integer
       
        Label_Name.Caption = Label_Name.Caption + CommonDialog1.FileTitle
        str_File = Me.CommonDialog1.FileName
        str_File = Left(str_File, Len(str_File) - Len(Me.CommonDialog1.FileTitle))
        Me.Label_Location.Caption = Me.Label_Location.Caption + str_File
        attr = GetFileAttributes(Me.CommonDialog1.FileName)
        If attr And FILE_ATTRIBUTE_ARCHIVE Then
            Me.Check_Archive.Value = 1
        Else
            Me.Check_Archive.Value = 0
        End If
        If attr And FILE_ATTRIBUTE_HIDDEN Then
            Me.Check_Hidden = 1
        Else
            Me.Check_Hidden.Value = 0
        End If
        If attr And FILE_ATTRIBUTE_READONLY Then
            Me.Check_ReadOnly = 1
        Else
            Me.Check_ReadOnly.Value = 0
        End If
        If attr And FILE_ATTRIBUTE_SYSTEM Then
            Me.Check_System = 1
        Else
            Me.Check_System.Value = 0
        End If
    End If
End Sub
